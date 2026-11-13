import pandas as pd
import os
from datetime import datetime
import requests  # Now used for API calls
# import pybaseball # Kept for completeness, but stubbed for now
# --- Configuration ---
MLB_START_YEAR = 2016
TODAY = datetime.now()
CURRENT_YEAR = TODAY.year 

# --- Directory Setup ---
PITCHERS_DIR = "raw data/pitchers"
HITTERS_DIR = "raw data/hitters"

def setup_directories():
    """Ensures output directories exist."""
    os.makedirs(PITCHERS_DIR, exist_ok=True)
    os.makedirs(HITTERS_DIR, exist_ok=True)
    print("Directories setup complete.")

# --- API SERVICE LAYER IMPLEMENTATION ---

# NOTE: These implementations are highly speculative based on general API patterns and the docs found. They may require refinement based on live API responses.

API_BASE_URL = "https://statsapi.mlb.com/api/v1/"
def _fetch_player_metadata(player_id: str, scope: str) -> dict:
    """
    Fetches biographical data and stats from the MLB Stats API for a player in a specific scope.
    This uses 'hydration' to request multiple data types in one call if the API supports it.
    """
    sport_id_map = {'MLB': 11, 'AA': 13, 'AAA': 12}
    sport_id = sport_id_map.get(scope)

    if not sport_id:
        print(f"[API ERROR] Unknown scope: {scope}")
        return {'name': f'Unknown Player {player_id}', 'position': 'N/A', 'age': None, 'handedness': 'N/A', 'team': 'N/A', 'stats': {}}

    # Endpoint to fetch person details and hydrate specific stats
    endpoint = f"people/{player_id}"
    params = {
        'sportId': sport_id,
        'hydrate': 'stats(group=hitting)', # Attempt to hydrate hitting stats
        'stats': 'season' # Explicitly request season stats in the params if hydration isn't enough
    }

    try:
        print(f"[StatsAPI] Querying metadata & stats for {player_id} ({scope})...")
        response = requests.get(f"{API_BASE_URL}{endpoint}", params=params)
        response.raise_for_status() # Raises HTTPError for bad responses (4xx or 5xx)
        data = response.json()

        # --- RESILIENT PARSING LOGIC (Adjusted for hydration response structure) ---
        # NOTE: Actual parsing depends heavily on the structure returned by the 'hydrate' parameter.

        # MOCK extraction for demonstration purposes, mirroring the original structure
        stats_container = data.get('stats', {}).get('season', {}).get('hitting', {}).get('stats', {})

        # Extracting key fields safely
        stats = stats_container.get('hits') or 0
        home_runs = stats_container.get('home_runs') or 0
        wrc = stats_container.get('wrc') or 0.0
        pi = stats_container.get('pi') or 0.0

        # MOCKing metadata extraction based on assumed successful call
        return {
            'name': f'Player_{player_id}_{scope}', # MOCK: Replace with actual name retrieval from data.person
            'position': 'R',  # MOCK: Replace with actual position retrieval from data.person
            'age': 30, # MOCK: Replace with actual age retrieval from data.person
            'handedness': 'R',
            'team': f'{scope} Team',
            'stats': {'hits': stats, 'home_runs': home_runs, 'wrc': wrc, 'pi': pi}
        }

    except requests.exceptions.HTTPError as e:
        print(f"[API HTTP ERROR] Received status {e.response.status_code} for {player_id} ({scope}). Check Player ID validity.")
        return {'name': f'Unknown Player {player_id}', 'position': 'N/A', 'age': None, 'handedness': 'N/A', 'team': 'N/A', 'stats': {}}
    except requests.exceptions.RequestException as e:
        print(f"[API CONNECTION ERROR] Could not connect to MLB Stats API for {player_id} ({scope}): {e}")
        return {'name': f'Unknown Player {player_id}', 'position': 'N/A', 'age': None, 'handedness': 'N/A', 'team': 'N/A', 'stats': {}}

def _fetch_stats_from_statsapi(player_id: str, scope: str) -> dict:
    """
    Fetches base stats from the MLB Stats API for a player in a specific scope.
    Delegates to the refined metadata/stats fetcher.
    """
    return _fetch_player_metadata(player_id, scope)

def _fetch_statcast_data_via_pybaseball(player_id: str, scope: str) -> dict:
    """
    Placeholder to call pybaseball/Savant for advanced metrics.
    """
    print(f"\n[pybaseball/Savant] Fetching Statcast data for {player_id} ({scope})...")
    # --- INTEGRATION POINT 2: IMPLEMENT PYBASEBALL CALL HERE ---
    # Example: stats = pybaseball.fastball_data.query(player_id, sportId=sport_id)

    # MOCK RETURN:
    return {'launch_angle': 18.5, 'exit_velocity': 95.2, 'spin_rate': 2200, 'astats': {'slugger_score': 1.2}}


# --- Core Logic Functions ---
def find_eligible_players() -> list[str]:
    """
    Step 1: Identifies all players who played in MLB between MLB_START_YEAR and CURRENT_YEAR.

    NOTE: This endpoint must be replaced with one that lists active players by season.
    """
    print(f"\n===================================================\nStep 1: Searching for MLB players from {MLB_START_YEAR} to {CURRENT_YEAR}...\n===================================================")
    # Mock list of player IDs found
    return ["player_123", "player_456"]

def aggregate_player_data(player_id: str) -> dict:
    """
    Step 2: Gathers all necessary data (MLB, AA, AAA) from multiple sources.
    """

    # Fetch Base Data (StatsAPI Spine)
    base_data = {
        'MLB': _fetch_stats_from_statsapi(player_id, 'MLB'),
        'AA': _fetch_stats_from_statsapi(player_id, 'AA'),
        'AAA': _fetch_stats_from_statsapi(player_id, 'AAA')
    }

    # Fetch Advanced Data (pybaseball/Savant Layer)
    # Note: Statcast data might not exist for all levels/seasons.
    statcast_mlb = _fetch_statcast_data_via_pybaseball(player_id, 'MLB')
    statcast_aaa = _fetch_statcast_data_via_pybaseball(player_id, 'AAA')

    # Combine all data into a single cohesive dictionary for further processing
    combined = {
        'name': base_data['MLB']['name'],
        'position': base_data['MLB']['position'],
        'age': base_data['MLB']['age'],
        'handedness': base_data['MLB']['handedness'],
        'team': base_data['MLB']['team'],
        'stats_mlb': base_data['MLB']['stats'],
        'stats_aa': base_data['AA']['stats'],
        'stats_aaa': base_data['AAA']['stats'],
        'statcast_mlb': statcast_mlb,
        'statcast_aaa': statcast_aaa
    }
    return combined

def process_and_save_player(player_data: dict, player_id: str):
    """
    Step 3: Determines player type, aggregates stats, and saves to XLSX.
    """
    position = player_data['position']

    # Determine player type
    is_hitter = position.upper() not in ['P']

    print(f"\n===================================================\nProcessing {player_data['name']} (Position: {position}). Saving to {'hitters' if is_hitter else 'pitchers'}...\n===================================================")

    # --- Aggregation Logic (Incorporating Multi-Source Data) ---
    final_stats = {}

    if is_hitter:
        # Combine all available hitting stats
        final_stats['Total_Hits'] = player_data['stats_mlb']['hits'] + player_data['stats_aa']['hits'] + player_data['stats_aaa']['hits']
        final_stats['MLB_HR'] = player_data['stats_mlb']['home_runs']
        # Incorporate Statcast data if available
        final_stats['MLB_Launch_Angle'] = player_data['statcast_mlb'].get('launch_angle', 'N/A')
        final_stats['MLB_EV'] = player_data['statcast_mlb'].get('exit_velocity', 'N/A')

        # Base Player Info
        final_stats['Name'] = player_data['name']
        final_stats['Age'] = player_data['age']
        final_stats['Handedness'] = player_data['handedness']

        output_dir = HITTERS_DIR
        filename = f"{player_id}_{player_data['name'].replace(' ', '_')}.xlsx"
    else: # Pitcher
        # Pitcher aggregation would happen here, pulling from relevant pitching endpoints
        final_stats['Name'] = player_data['name']
        final_stats['Age'] = player_data['age']
        final_stats['Handedness'] = player_data['handedness']

        output_dir = PITCHERS_DIR
        filename = f"{player_id}_{player_data['name'].replace(' ', '_')}.xlsx"

    # --- Saving to XLSX ---
    df = pd.DataFrame([final_stats])
    save_path = os.path.join(output_dir, filename)
    df.to_excel(save_path, index=False)
    print(f"\nSUCCESS: Saved data for {player_data['name']} to {save_path}")


def main():
    setup_directories()

    # 1. Identify target players
    eligible_ids = find_eligible_players()

    if not eligible_ids:
        print("Exiting: No eligible players found based on MLB 2016+ criteria.")
        return

    print(f"\n===================================================\nStarting data retrieval and processing for {len(eligible_ids)} players.\n===================================================")

    # --- PERFORMANCE OPTIMIZATION: PARALLEL EXECUTION ---
    # Using ThreadPoolExecutor to run network-bound tasks concurrently (I/O Bound)
    from concurrent.futures import ThreadPoolExecutor, as_completed

    # Max workers set to a reasonable number based on available cores and API limits.
    MAX_WORKERS = 20

    all_results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        # Submit all tasks to the executor
        future_to_player = {
            executor.submit(aggregate_player_data, player_id): player_id
            for player_id in eligible_ids
        }

        print(f"Submitting {len(eligible_ids)} player data requests to ThreadPoolExecutor with {MAX_WORKERS} workers...")

        # Wait for results as they complete
        for future in as_completed(future_to_player):
            player_id = future_to_player[future]
            try:
                player_data = future.result()

                # Process and save *after* data aggregation is complete for that thread
                process_and_save_player(player_data, player_id)
            except Exception as exc:
                print(f"\n!!! EXCEPTION encountered for player {player_id} during aggregation: {exc} !!!")


if __name__ == "__main__":
    main()