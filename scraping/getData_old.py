import sys
import sqlite3
import pandas as pd
import statsapi as sa
import requests
import time
import datetime

from pathlib import Path

# ==================================================================================================
# set up
# ==================================================================================================
HERE = Path(__file__).resolve().parent         # .../scraping
PROJECT_ROOT = HERE.parent                     # .../baseballScouting

sys.path.insert(0, str(PROJECT_ROOT))
currentYear = datetime.year()
startYear = 2016

# ==================================================================================================
# load data
# ==================================================================================================
def loadTables(db_path, table_names):
    """Load the first matching table from a SQLite DB; return empty DataFrame if unavailable."""
    if isinstance(table_names, str):
        table_names = [table_names]

    if not db_path.exists() or db_path.stat().st_size == 0:
        return pd.DataFrame()

    conn = sqlite3.connect(db_path)
    try:
        available_tables = {
            row[0]
            for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
        }

        for table_name in table_names:
            if table_name in available_tables:
                return pd.read_sql_query(f"select * from {table_name}", conn)

        return pd.DataFrame()
    finally:
        conn.close()

hitting_df = {}
for year in range(startYear, currentYear):
    db_path = PROJECT_ROOT / 'clean data' / 'league hitting' / f'league_hitting_{year}.db'
    if not db_path.exists():
        db_path = PROJECT_ROOT / 'clean data' / 'league hitting' / f'hitting_{year}.db'
    hitting_df[year] = loadTables(db_path, [f'hitting_{year}', f'league_hitting_{year}'])
    
pitching_df = {}
for year in range(startYear, currentYear):
    db_path = PROJECT_ROOT / 'clean data' / 'league pitching' / f'league_pitching_{year}.db'
    if not db_path.exists():
        db_path = PROJECT_ROOT / 'clean data' / 'league pitching' / f'pitching_{year}.db'
    pitching_df[year] = loadTables(db_path, [f'pitching_{year}', f'league_pitching_{year}'])

# ==================================================================================================
# functions
# ==================================================================================================

def get_birth_date(player_id):
    """Fetch player's birth date from statsapi."""
    try:
        url = f"https://statsapi.mlb.com/api/v1/people/{player_id}"
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json()
        people = data.get('people', [])
        if people:
            birth_date_str = people[0].get('birthDate')
            if birth_date_str:
                return pd.to_datetime(birth_date_str)
    except requests.RequestException as exc:
        print(f"  Warning: Could not fetch birth date for player {player_id}: {exc}")
    return None


def calculate_age(birth_date, season):
    """Calculate age during a given season (age at July 1 of that season)."""
    if birth_date is None or season is None:
        return None
    try:
        # Extract year from season (handle formats like '2018' or '2018.1')
        year = int(float(season))
        season_date = pd.Timestamp(year=year, month=7, day=1)
        age = (season_date - birth_date).days // 365.25
        return int(age)
    except (ValueError, TypeError):
        return None


def normalize_level_to_key(level_value):
    """Normalize level labels to a shared key for cross-source merges."""
    if pd.isna(level_value):
        return None

    text = str(level_value).strip().lower()
    mapping = {
        'mlb': 'MLB',
        'aaa': 'AAA',
        'aa': 'AA',
        'a+': 'A+',
        'high-a': 'A+',
        'single-a': 'A',
        'a': 'A',
        'a-': 'A-',
        'short-season a': 'A-',
        'r': 'R',
        'rookie': 'R',
        'rookie (adv)': 'R',
        'milb': 'MiLB',
        'ros': 'ROS',
    }
    return mapping.get(text)


def add_prefixed_fields(row, data, prefix):
    """Copy a dict's keys into a row with a prefix, preserving every raw API field."""
    for key, value in data.items():
        row[f"{prefix}{key}"] = value


def enrich_hitting_with_fangraphs(career_df, player_name, fg_dir=None):
    """Merge FanGraphs-only hitting fields from local per-player CSV exports."""
    if career_df.empty:
        return career_df

    fg_dir = fg_dir or FANGRAPHS_HITTING_DIR
    fg_path = fg_dir / f"{player_name}.csv"
    if not fg_path.exists():
        return career_df

    try:
        fg_df = pd.read_csv(fg_path)
    except Exception as exc:
        print(f"  Warning: Could not read FanGraphs CSV for {player_name}: {exc}")
        return career_df

    if fg_df.empty:
        return career_df

    required = {'Season', 'Level'}
    if not required.issubset(set(fg_df.columns)):
        return career_df

    fg_cols = [
        'wRC', 'wRAA', 'wOBA', 'wRC+',
        'IFH%', 'BUH%', 'Pull%', 'Cent%', 'Oppo%', 'Soft%', 'Med%', 'Hard%',
        'Spd', 'UBR', 'wGDP', 'XBR', 'wSB',
    ]
    available_fg_cols = [col for col in fg_cols if col in fg_df.columns]
    if not available_fg_cols:
        return career_df

    fg_df = fg_df.copy()
    fg_df['season_key'] = pd.to_numeric(fg_df['Season'], errors='coerce').astype('Int64')
    fg_df['level_key'] = fg_df['Level'].apply(normalize_level_to_key)
    fg_df = fg_df[fg_df['season_key'].notna() & fg_df['level_key'].notna()]
    if fg_df.empty:
        return career_df

    fg_merge = fg_df[['season_key', 'level_key'] + available_fg_cols].copy()
    fg_merge = fg_merge.drop_duplicates(subset=['season_key', 'level_key'], keep='first')

    out = career_df.copy()
    out['season_key'] = pd.to_numeric(out['season'], errors='coerce').astype('Int64')
    out['level_key'] = out['level'].apply(normalize_level_to_key)
    out = out.merge(fg_merge, how='left', on=['season_key', 'level_key'], suffixes=('', '_fg'))

    for col in available_fg_cols:
        fg_col = f"{col}_fg"
        if fg_col in out.columns:
            if col not in out.columns:
                out[col] = out[fg_col]
            else:
                out[col] = out[col].combine_first(out[fg_col])
            out = out.drop(columns=[fg_col])

    return out.drop(columns=['season_key', 'level_key'])


def get_full_career(player_id, player_name, group='hitting'):
    """
    Fetch full career stats (all levels) for a player.
    
    Args:
        player_id: MLB Advanced Media ID (mlbam_id)
        player_name: Player name for logging
        group: 'hitting' or 'pitching'
    
    Returns:
        DataFrame with career stats across all levels
    """
    # Fetch birth date for age calculation
    birth_date = get_birth_date(player_id)
    
    sport_ids = {
        1:  'MLB',
        11: 'AAA',
        12: 'AA',
        13: 'High-A',
        14: 'Single-A'
    }

    all_rows = []

    for sport_id, label in sport_ids.items():
        url = f"https://statsapi.mlb.com/api/v1/people/{player_id}/stats"
        params = {
            "stats": "yearByYear,yearByYearAdvanced",
            "group": group,
            "sportId": sport_id,
        }
        try:
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()
            data = response.json()
        except Exception as e:
            print(f"  Warning: Error fetching {group} data for {player_name} (ID: {player_id}): {e}")
            continue

        stats_list = data.get('stats', [])
        if not stats_list:
            continue

        regular_splits = []
        advanced_splits = []
        for stat_block in stats_list:
            stat_type = stat_block.get('type', {}).get('displayName', '')
            block_splits = stat_block.get('splits', [])
            if stat_type == 'yearByYear':
                regular_splits = block_splits
            elif stat_type == 'yearByYearAdvanced':
                advanced_splits = block_splits

        if not regular_splits:
            continue

        # Build lookup for advanced hitting stats on the same seasonal stint.
        # Keying by season/team/league keeps multiple stints distinct.
        advanced_lookup = {}
        if group == 'hitting' or group == 'pitching':
            for adv_split in advanced_splits:
                adv_key = (
                    adv_split.get('season'),
                    adv_split.get('team', {}).get('id'),
                    adv_split.get('league', {}).get('id'),
                )
                advanced_lookup[adv_key] = adv_split.get('stat', {})

        for split in regular_splits:
            s = split['stat']
            season = split.get('season')
            age = calculate_age(birth_date, season)
            split_key = (
                season,
                split.get('team', {}).get('id'),
                split.get('league', {}).get('id'),
            )
            adv = advanced_lookup.get(split_key, {})
            
            if group == 'hitting':
                ground_events = (adv.get('groundOuts') or 0) + (adv.get('groundHits') or 0)
                line_events = (adv.get('lineOuts') or 0) + (adv.get('lineHits') or 0)
                fly_events = (adv.get('flyOuts') or 0) + (adv.get('flyHits') or 0)
                pop_events = (adv.get('popOuts') or 0) + (adv.get('popHits') or 0)
                total_batted_events = ground_events + line_events + fly_events + pop_events

                ld_pct = (line_events / total_batted_events) if total_batted_events else None
                gb_pct = (ground_events / total_batted_events) if total_batted_events else None
                fb_pct = ((fly_events + pop_events) / total_batted_events) if total_batted_events else None
                gb_fb = (ground_events / (fly_events + pop_events)) if (fly_events + pop_events) else None
                # Approximation: StatsAPI exposes pop balls but not explicit infield-fly splits.
                iffb_pct = (pop_events / (fly_events + pop_events)) if (fly_events + pop_events) else None
                hr_fb = (float(s.get('homeRuns') or 0) / (fly_events + pop_events)) if (fly_events + pop_events) else None

                all_rows.append({
                    'mlbam_id': player_id,
                    'player_name': player_name,
                    'season': season,
                    'age': age,
                    'level':  label,
                    'team':   split.get('team', {}).get('name'),
                    'league': split.get('league', {}).get('name', ''),
                    'G':      s.get('gamesPlayed'),
                    'PA':     s.get('plateAppearances'),
                    'AB':     s.get('atBats'),
                    'R':      s.get('runs'),
                    'H':      s.get('hits'),
                    '2B':     s.get('doubles'),
                    '3B':     s.get('triples'),
                    'HR':     s.get('homeRuns'),
                    'RBI':    s.get('rbi'),
                    'TB':     s.get('totalBases'),
                    'SB':     s.get('stolenBases'),
                    'CS':     s.get('caughtStealing'),
                    'BB':     s.get('baseOnBalls'),
                    'IBB':    s.get('intentionalWalks'),
                    'HBP':    s.get('hitByPitch'),
                    'SH':     s.get('sacBunts'),
                    'SF':     s.get('sacFlies'),
                    'GIDP':   s.get('groundIntoDoublePlay'),
                    'SO':     s.get('strikeOuts'),
                    'BA':     s.get('avg'),
                    'AVG':    s.get('avg'),
                    'OBP':    s.get('obp'),
                    'SLG':    s.get('slg'),
                    'OPS':    s.get('ops'),
                    # Advanced fields available from StatsAPI (closest overlap to league table)
                    'ISO':    adv.get('iso'),
                    'BAbip':  adv.get('babip'),
                    'BB%':    adv.get('walksPerPlateAppearance'),
                    'SO%':    adv.get('strikeoutsPerPlateAppearance'),
                    'BB/K':   adv.get('walksPerStrikeout'),
                    'HR%':    adv.get('homeRunsPerPlateAppearance'),
                    # Additional StatsAPI advanced fields
                    'BIP':    adv.get('ballsInPlay'),
                    'Pitches': adv.get('numberOfPitches'),
                    'P/PA':   adv.get('pitchesPerPlateAppearance'),
                    'SwStr':  adv.get('swingAndMisses'),
                    'Swings': adv.get('totalSwings'),
                    'GO_adv': adv.get('groundOuts'),
                    'FO_adv': adv.get('flyOuts'),
                    'LO_adv': adv.get('lineOuts'),
                    'PO_adv': adv.get('popOuts'),
                    'GH_adv': adv.get('groundHits'),
                    'FH_adv': adv.get('flyHits'),
                    'LH_adv': adv.get('lineHits'),
                    'PH_adv': adv.get('popHits'),
                    'ROE':    adv.get('reachedOnError'),
                    # Approximate batted-ball profile from available event buckets
                    'LD%':    ld_pct,
                    'GB%':    gb_pct,
                    'FB%':    fb_pct,
                    'IFFB%':  iffb_pct,
                    'GB/FB':  gb_fb,
                    'HR/FB':  hr_fb
                })
            elif group == 'pitching':
                row = {
                    'mlbam_id': player_id,
                    'player_name': player_name,
                    'season': season,
                    'age': age,
                    'level':  label,
                    'team':   split.get('team', {}).get('name'),
                    'league': split.get('league', {}).get('name', ''),
                    'G':      s.get('gamesPlayed'),
                    'IP':     s.get('inningsPitched'),
                    'W':      s.get('wins'),
                    'L':      s.get('losses'),
                    'SV':     s.get('saves'),
                    'SVO':    s.get('saveOpportunities'),
                    'HLD':    s.get('holds'),
                    'BS':     s.get('blownSaves'),
                    'ERA':    s.get('era'),
                    'H':      s.get('hits'),
                    'R':      s.get('runs'),
                    'ER':     s.get('earnedRuns'),
                    'HR':     s.get('homeRuns'),
                    'BB':     s.get('baseOnBalls'),
                    'SO':     s.get('strikeOuts'),
                    'HBP':    s.get('hitBatsmen'),
                    'WHIP':   s.get('whip'),
                    'BAA':    s.get('avg'),
                    'CS':     s.get('caughtStealing'),
                    'SB':     s.get('stolenBases'),
                    'BLK':    s.get('balks'),
                    'WP':     s.get('wildPitches'),
                    'GO':     s.get('groundOuts'),
                    'AO':     s.get('airOuts'),
                }
                add_prefixed_fields(row, s, 'api_')
                add_prefixed_fields(row, adv, 'adv_')
                all_rows.append(row)

    if not all_rows:
        return pd.DataFrame()

    df = pd.DataFrame(all_rows)

    # Clean up numeric columns based on group
    if group == 'hitting':
        num_cols = ['G', 'PA', 'AB', 'R', 'H', '2B', '3B', 'HR', 'RBI', 'TB', 'SB', 'CS', 'BB', 'IBB', 'HBP', 'SH', 'SF', 'GIDP', 'SO', 'BIP', 'Pitches', 'SwStr', 'Swings', 'GO_adv', 'FO_adv', 'LO_adv', 'PO_adv', 'GH_adv', 'FH_adv', 'LH_adv', 'PH_adv', 'ROE']
    else:  # pitching
        num_cols = ['G', 'W', 'L', 'SV', 'SVO', 'HLD', 'BS', 'H', 'R', 'ER', 'HR', 'BB', 'SO', 'HBP', 'CS', 'SB', 'BLK', 'WP', 'GO', 'AO']
    
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').astype('Int64')

    if group == 'hitting':
        pct_cols = ['BB%', 'SO%', 'HR%', 'BB/K', 'LD%', 'GB%', 'FB%']
        decimal_cols = ['BA', 'AVG', 'OBP', 'SLG', 'OPS', 'ISO', 'BAbip', 'GB/FB', 'P/PA']
        for col in pct_cols + decimal_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

    # Sort by season then by AB/IP descending (so full seasons appear before rehab stints)
    sort_col = 'AB' if group == 'hitting' else 'IP'
    if sort_col in df.columns:
        df = df.sort_values(['season', sort_col], ascending=[True, False]).reset_index(drop=True)

    return df


def fetch_and_save_all_players(hitting_df, pitching_df, group='hitting', db_dir=None, table_name=None):
    """
    Fetch career stats for all unique players and save each to individual SQL database.
    
    Args:
        hitting_df: Dictionary of hitting DataFrames by year
        pitching_df: Dictionary of pitching DataFrames by year
        group: 'hitting' or 'pitching'
        db_dir: Directory where individual player databases will be saved
        table_name: Table name in each database
    """
    # Extract unique players from the appropriate data
    df_dict = hitting_df if group == 'hitting' else pitching_df
    unique_players = {}  # mlbam_id -> player_name
    
    for year_data in df_dict.values():
        if 'mlbam_id' in year_data.columns and 'Player' in year_data.columns:
            for _, row in year_data.iterrows():
                pid = row['mlbam_id']
                if pd.notna(pid):
                    unique_players[int(pid)] = row['Player']
    
    print(f"\nFound {len(unique_players)} unique players in {group} data")
    
    # Create directory if needed
    if db_dir and table_name:
        db_dir.mkdir(parents=True, exist_ok=True)
    
    total_rows = 0
    for idx, (player_id, player_name) in enumerate(unique_players.items(), 1):
        print(f"  [{idx}/{len(unique_players)}] Fetching {player_name} ({group})...", end="", flush=True)
        career_df = get_full_career(player_id, player_name, group=group)

        if group == 'hitting' and not career_df.empty:
            career_df = enrich_hitting_with_fangraphs(career_df, player_name)
        
        if not career_df.empty:
            career_df = career_df.round(3)

            # Save to individual player database
            if db_dir and table_name:
                # Use full player name as filename
                player_db_path = db_dir / f"{player_name}.db"
                conn = sqlite3.connect(player_db_path)
                career_df.to_sql(table_name, conn, if_exists='replace', index=False)
                conn.close()
                total_rows += len(career_df)
                print(f" ({len(career_df)} rows)", flush=True)
            else:
                print(f" ({len(career_df)} rows)", flush=True)
        else:
            print(" (no data)", flush=True)
        
        # Be respectful to the API
        time.sleep(0.5)
    
    print(f"  Total rows saved: {total_rows}")

# ==================================================================================================
# scraping data
# ==================================================================================================

if __name__ == '__main__':
    # Fetch hitting data
    print("=" * 80)
    print("FETCHING HITTING DATA")
    print("=" * 80)
    hitting_db_dir = PROJECT_ROOT / 'raw data' / 'player hitting'
    fetch_and_save_all_players(
        hitting_df, pitching_df, 
        group='hitting',
        db_dir=hitting_db_dir,
        table_name='player_hitting_career'
    )
    
    # Fetch pitching data
    print("\n" + "=" * 80)
    print("FETCHING PITCHING DATA")
    print("=" * 80)
    pitching_db_dir = PROJECT_ROOT / 'raw data' / 'player pitching'
    fetch_and_save_all_players(
        hitting_df, pitching_df, 
        group='pitching',
        db_dir=pitching_db_dir,
        table_name='player_pitching_career'
    )
    
    print("\n" + "=" * 80)
    print("COMPLETE")
    print("=" * 80)