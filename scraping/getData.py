"""
MLB scouting data collector (multi-source).

One XLSX per player with season-level data only (no pitch-by-pitch):
  Bio | Hitting | Pitching | Fielding | Awards | Draft | BRef_Hitting_WAR | BRef_Pitching_WAR

Sources
  - MLB Stats API (statsapi.mlb.com): AA/AAA/MLB year-by-year hitting, pitching, fielding, awards, draft, bio
  - Baseball-Reference bulk WAR files: WAR, OPS+, ERA+ (MLB only)
  - Baseball Savant leaderboards: season-level Statcast (MLB only, 2015+)
  - FanGraphs JSON leaderboards: advanced stats (MLB only, best effort - may be blocked)
  - Chadwick Bureau register: ID cross-reference (BBRef / FanGraphs / Retrosheet IDs)

Every source is optional. If one fails, the script says so and carries on with the rest.
All downloads are cached in "raw data/_cache" so re-runs are fast (the current season is always re-fetched).
"""
import io
import json
import os
import re
import time
from datetime import date, datetime
from functools import reduce

import warnings

import pandas as pd
import requests

warnings.filterwarnings("ignore", category=pd.errors.PerformanceWarning)

try:                                   # xlsxwriter is much faster than openpyxl for wide sheets
    import xlsxwriter  # noqa: F401
    EXCEL_ENGINE = "xlsxwriter"
except ImportError:
    EXCEL_ENGINE = "openpyxl"

# =============================== Configuration ===============================
MLB_START_YEAR = 2016          # ONLY decides which players are included (played in MLB 2016+)
CURRENT_YEAR = datetime.now().year

SOURCES = {                    # flip any to False to skip a source
    "bref": True,
    "savant": True,
    "fangraphs": True,
    "draft": True,
    "awards": True,
    "chadwick": True,
}
SAVANT_START_YEAR = 2015       # Statcast era
FANGRAPHS_START_YEAR = 2005
DRAFT_START_YEAR = 1965

BASE_URL = "https://statsapi.mlb.com/api/v1"
MAX_RETRIES = 4
BATCH_SIZE = 50                # player IDs per /people request
PAGE_SIZE = 1000
REQUEST_DELAY = 0.2            # MLB API
SAVANT_DELAY = 1.0
FANGRAPHS_DELAY = 3.0          # be polite to FanGraphs

LEVELS = {1: "MLB", 11: "AAA", 12: "AA"}
LEVEL_RANK = {"AA": 0, "AAA": 1, "MLB": 2}

BREF_BAT_URL = "https://www.baseball-reference.com/data/war_daily_bat.txt"
BREF_PIT_URL = "https://www.baseball-reference.com/data/war_daily_pitch.txt"
CHADWICK_URL = "https://raw.githubusercontent.com/chadwickbureau/register/master/data/people-{}.csv"

CACHE_DIR = "raw data/_cache"
PITCHERS_DIR = "raw data/pitchers"
HITTERS_DIR = "raw data/hitters"

BROWSER_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")

META_COLS = ["Player", "MLBAM_ID", "Year", "Age_Apr1", "Age_Apr1_Decimal", "Level", "Team",
             "Team_ID", "League", "Position", "Is_Total_Row", "Is_Synthetic_Total", "Num_Teams"]
PLACEHOLDERS = ["-.--", ".---", "*.**", "-", "--", "---", "inf", "Infinity", "-inf"]

# Savant leaderboards (URLs from pybaseball's source). (group, key, url, pivot_by_pitch_type)
_SV = "https://baseballsavant.mlb.com"
SAVANT_DATASETS = [
    ("hit", "exp", _SV + "/leaderboard/expected_statistics?type=batter&year={y}&position=&team=&min=1&csv=true", False),
    ("hit", "ev", _SV + "/leaderboard/statcast?type=batter&year={y}&position=&team=&min=1&csv=true", False),
    ("hit", "pct", _SV + "/leaderboard/percentile-rankings?type=batter&year={y}&position=&team=&csv=true", False),
    ("hit", "sprint", _SV + "/leaderboard/sprint_speed?year={y}&position=&team=&min=1&csv=true", False),
    ("hit", "vspitch", _SV + "/leaderboard/pitch-arsenal-stats?type=batter&pitchType=&year={y}&team=&min=1&csv=true", True),
    ("hit", "oaa", _SV + "/leaderboard/outs_above_average?type=Fielder&year={y}&team=&range=year&min=1&pos=&roles=&viz=show&csv=true", False),
    ("hit", "framing", _SV + "/catcher_framing?year={y}&team=&min=1&sort=4,1&csv=true", False),
    ("hit", "poptime", _SV + "/leaderboard/poptime?year={y}&team=&min2b=1&min3b=0&csv=true", False),
    ("hit", "ofjump", _SV + "/leaderboard/outfield_jump?year={y}&min=1&csv=true", False),
    ("pit", "exp", _SV + "/leaderboard/expected_statistics?type=pitcher&year={y}&position=&team=&min=1&csv=true", False),
    ("pit", "ev", _SV + "/leaderboard/statcast?type=pitcher&year={y}&position=&team=&min=1&csv=true", False),
    ("pit", "pct", _SV + "/leaderboard/percentile-rankings?type=pitcher&year={y}&position=&team=&csv=true", False),
    ("pit", "arsenal_speed", _SV + "/leaderboard/pitch-arsenals?year={y}&min=1&type=avg_speed&hand=&csv=true", False),
    ("pit", "arsenal_spin", _SV + "/leaderboard/pitch-arsenals?year={y}&min=1&type=avg_spin&hand=&csv=true", False),
    ("pit", "arsenal_usage", _SV + "/leaderboard/pitch-arsenals?year={y}&min=1&type=n_&hand=&csv=true", False),
    ("pit", "arsenal_stats", _SV + "/leaderboard/pitch-arsenal-stats?type=pitcher&pitchType=&year={y}&team=&min=1&csv=true", True),
]
SAVANT_ID_COLS = ["player_id", "pitcher", "batter", "mlb_id", "key_mlbam"]
SAVANT_DROP_COLS = {"last_name, first_name", "player_name", "last_name", "first_name", "name", "entity_name",
                    "team_name", "team_name_alt", "team", "year", "entity_id"}

# FanGraphs leaderboard "type" ids (Dashboard, Advanced, Batted Ball, Win Prob, Value, Plate Disc, Pitch Type, Pitch Value)
FANGRAPHS_TYPES = [8, 1, 2, 3, 4, 5, 6, 7]

session = requests.Session()
session.headers.update({"User-Agent": "baseball-scouting-project/1.0"})

USE_ADVANCED = True            # downgrade flags, flipped automatically if the API rejects a hydrate
USE_FIELDING = True
LAST_STATUS = None


def setup_directories():
    for d in (PITCHERS_DIR, HITTERS_DIR, CACHE_DIR):
        os.makedirs(d, exist_ok=True)
    print(f"Directories setup complete. (Excel engine: {EXCEL_ENGINE})")


# =============================== HTTP + cache ===============================

def api_get(path: str, params: dict | None = None) -> dict | None:
    """MLB Stats API GET with retries. Fails fast on 4xx (except 429); sets LAST_STATUS."""
    global LAST_STATUS
    url = f"{BASE_URL}{path}"
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.get(url, params=params or {}, timeout=60)
            if resp.status_code == 429:
                wait = int(resp.headers.get("Retry-After", 5 * attempt))
                print(f"[WARNING] 429 rate limited, sleeping {wait}s (attempt {attempt}/{MAX_RETRIES})")
                time.sleep(wait)
                continue
            if 400 <= resp.status_code < 500:
                LAST_STATUS = resp.status_code
                print(f"[WARNING] {path} returned {resp.status_code}; not retrying.")
                return None
            resp.raise_for_status()
            time.sleep(REQUEST_DELAY)
            return resp.json()
        except Exception as e:
            print(f"[WARNING] {path} attempt {attempt}/{MAX_RETRIES} failed: {e}")
            time.sleep(2 * attempt)
    LAST_STATUS = 599
    print(f"[CRITICAL] Giving up on {path}")
    return None


def http_text(url: str, delay: float = 0.5, timeout: int = 120, max_retries: int = 3,
              headers: dict | None = None) -> str | None:
    """GET a non-MLB URL as text. Returns None if blocked/failed (never raises)."""
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.get(url, timeout=timeout, headers=headers or {"User-Agent": BROWSER_UA})
            if resp.status_code == 429:
                wait = int(resp.headers.get("Retry-After", 10 * attempt))
                print(f"[WARNING] 429 from {url.split('?')[0]}; sleeping {wait}s")
                time.sleep(wait)
                continue
            if resp.status_code >= 400 and resp.status_code < 500:
                print(f"[WARNING] {resp.status_code} from {url.split('?')[0]}")
                return None
            resp.raise_for_status()
            time.sleep(delay)
            return resp.text
        except Exception as e:
            print(f"[WARNING] {url.split('?')[0]} attempt {attempt}/{max_retries} failed: {e}")
            time.sleep(2 * attempt)
    return None


def _cache_file(name: str) -> str:
    return os.path.join(CACHE_DIR, name)


def load_cached(name: str, year: int | None = None):
    """Cached DataFrame, or None. The current season is never served from cache."""
    if year is not None and year >= CURRENT_YEAR:
        return None
    path = _cache_file(name)
    if os.path.exists(path):
        try:
            return pd.read_pickle(path)
        except Exception:
            return None
    return None


def save_cached(name: str, df: pd.DataFrame):
    try:
        df.to_pickle(_cache_file(name))
    except Exception as e:
        print(f"[WARNING] could not cache {name}: {e}")


def index_by_player(df: pd.DataFrame | None) -> dict[int, pd.DataFrame]:
    if df is None or df.empty:
        return {}
    return {int(k): g.drop(columns=["mlbam"]).reset_index(drop=True) for k, g in df.groupby("mlbam")}


def merge_new_cols(base: pd.DataFrame, other: pd.DataFrame, key: str = "mlbam") -> pd.DataFrame:
    """Outer-merge only the columns of `other` that `base` doesn't already have."""
    new = [c for c in other.columns if c not in base.columns]
    if not new:
        return base
    return base.merge(other[[key] + new], on=key, how="outer")


# =============================== Discovery ===============================

def discover_player_ids(start_year: int, end_year: int) -> list[int]:
    print(f"[Discovery] Finding MLB players from {start_year} through {end_year}...")
    ids: set[int] = set()
    for year in range(start_year, end_year + 1):
        year_ids: set[int] = set()
        for group in ("hitting", "pitching"):
            offset = 0
            while True:
                data = api_get("/stats", {
                    "stats": "season", "group": group, "season": year, "sportIds": 1,
                    "playerPool": "All", "limit": PAGE_SIZE, "offset": offset,
                })
                if not data or not data.get("stats"):
                    break
                splits = data["stats"][0].get("splits", [])
                for s in splits:
                    pid = s.get("player", {}).get("id")
                    if pid:
                        year_ids.add(pid)
                if len(splits) < PAGE_SIZE:
                    break
                offset += PAGE_SIZE
        print(f"-> {year}: {len(year_ids)} players")
        ids |= year_ids
    print(f"[Discovery] {len(ids)} unique players total.")
    return sorted(ids)


# =============================== Source: Baseball-Reference ===============================

def load_bref(url: str, cache_name: str) -> dict[int, pd.DataFrame]:
    """BRef bulk WAR file -> {mlbam_id: rows}. Empty dict on failure."""
    cache_path = _cache_file(cache_name)
    try:
        if os.path.exists(cache_path):
            print(f"[BRef] Using cached {cache_name}")
            text = open(cache_path, encoding="utf-8").read()
        else:
            print(f"[BRef] Downloading {url} ...")
            text = http_text(url, delay=1.0)
            if not text or text.lstrip().startswith("<"):
                raise RuntimeError("no usable response (blocked or HTML page)")
            with open(cache_path, "w", encoding="utf-8") as f:
                f.write(text)
        df = pd.read_csv(io.StringIO(text), low_memory=False)
        df["mlb_ID"] = pd.to_numeric(df["mlb_ID"], errors="coerce")
        df = df.dropna(subset=["mlb_ID"])
        df["mlb_ID"] = df["mlb_ID"].astype(int)
        print(f"[BRef] {cache_name}: {len(df)} rows, {df['mlb_ID'].nunique()} players")
        return {k: g.drop(columns=["mlb_ID"]).reset_index(drop=True) for k, g in df.groupby("mlb_ID")}
    except Exception as e:
        print(f"[BRef WARNING] Could not load {cache_name}: {e}. Continuing without it.")
        return {}


def aggregate_bref(raw: pd.DataFrame | None, metric: str, weight: str) -> pd.DataFrame | None:
    """One row per year: BRef_WAR (summed over stints) and BRef_<metric> (PA/IP-weighted if >1 stint)."""
    if raw is None or raw.empty or "WAR" not in raw.columns or "year_ID" not in raw.columns:
        return None
    rows = []
    for year, st in raw.groupby("year_ID"):
        war = pd.to_numeric(st["WAR"], errors="coerce")
        rec = {"Year": int(year), "BRef_WAR": round(war.sum(), 2) if war.notna().any() else None}
        if metric in st.columns:
            vals = pd.to_numeric(st[metric], errors="coerce")
            if len(st) == 1:
                rec[f"BRef_{metric}"] = vals.iloc[0]
            elif weight in st.columns:
                w = pd.to_numeric(st[weight], errors="coerce")
                ok = vals.notna() & w.notna() & (w > 0)
                rec[f"BRef_{metric}"] = round((vals[ok] * w[ok]).sum() / w[ok].sum(), 1) if ok.any() else None
        rows.append(rec)
    return pd.DataFrame(rows)


# =============================== Source: Baseball Savant ===============================

def parse_savant(text: str, key: str, pivot: bool) -> pd.DataFrame | None:
    """Parse one Savant CSV. None = unusable response (don't cache); empty DataFrame = valid but no rows."""
    t = text.lstrip("\ufeff").strip()
    if not t or t.startswith("<"):
        return None
    try:
        df = pd.read_csv(io.StringIO(t))
    except Exception:
        return None
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
    if df.empty:
        return pd.DataFrame(columns=["mlbam"])
    idcol = next((c for c in SAVANT_ID_COLS if c in df.columns), None)
    if idcol is None:
        return None
    df = df.rename(columns={idcol: "mlbam"})
    df = df.drop(columns=[c for c in df.columns if c in SAVANT_DROP_COLS])
    df["mlbam"] = pd.to_numeric(df["mlbam"], errors="coerce")
    df = df.dropna(subset=["mlbam"])
    df["mlbam"] = df["mlbam"].astype(int)

    if pivot and "pitch_type" in df.columns:
        values = [c for c in df.columns if c not in ("mlbam", "pitch_type", "pitch_name")]
        wide = df.pivot_table(index="mlbam", columns="pitch_type", values=values, aggfunc="first")
        wide.columns = [f"{pt}_{col}" for col, pt in wide.columns]
        df = wide.reset_index()
    else:
        df = df.drop_duplicates("mlbam")
    return df.rename(columns={c: f"SV_{key}_{c}" for c in df.columns if c != "mlbam"})


def savant_dataset_year(group: str, key: str, url: str, pivot: bool, year: int) -> pd.DataFrame | None:
    cname = f"savant_{group}_{key}_{year}.pkl"
    cached = load_cached(cname, year)
    if cached is not None:
        return cached
    text = http_text(url.format(y=year), delay=SAVANT_DELAY)
    if text is None:
        return None
    df = parse_savant(text, key, pivot)
    if df is not None:
        save_cached(cname, df)
    return df


def load_savant() -> dict[str, pd.DataFrame]:
    print(f"\n[Savant] Season-level Statcast {SAVANT_START_YEAR}-{CURRENT_YEAR} "
          f"({len(SAVANT_DATASETS)} leaderboards per year)...")
    frames: dict[str, list] = {"hit": [], "pit": []}
    ok = fail = 0
    for year in range(SAVANT_START_YEAR, CURRENT_YEAR + 1):
        per_group: dict[str, list] = {"hit": [], "pit": []}
        for group, key, url, pivot in SAVANT_DATASETS:
            df = savant_dataset_year(group, key, url, pivot, year)
            if df is None:
                fail += 1
                if ok == 0 and fail >= 6:
                    print("[Savant WARNING] Every request failed - Savant looks blocked. Skipping Savant.")
                    return {"hit": pd.DataFrame(), "pit": pd.DataFrame()}
                continue
            ok += 1
            if not df.empty:
                per_group[group].append(df)
        counts = []
        for g, parts in per_group.items():
            if parts:
                wide = reduce(lambda a, b: a.merge(b, on="mlbam", how="outer"), parts)
                wide["Year"] = year
                frames[g].append(wide)
                counts.append(f"{g}: {len(wide)} players, {wide.shape[1] - 2} cols")
        print(f"-> Savant {year}: " + (" | ".join(counts) if counts else "no data"))
    return {g: (pd.concat(f, ignore_index=True) if f else pd.DataFrame()) for g, f in frames.items()}


# =============================== Source: FanGraphs (best effort) ===============================

def _fg_url(stats: str, season: int, type_id: int) -> str:
    return ("https://www.fangraphs.com/api/leaders/major-league/data"
            f"?age=&pos=all&stats={stats}&lg=all&qual=0&season={season}&season1={season}"
            "&startdate=&enddate=&month=0&hand=&team=0&pageitems=2000000000&pagenum=1&ind=0"
            f"&rost=0&players=0&type={type_id}&postseason=&sortdir=default&sortstat=WAR")


def parse_fangraphs(text: str) -> pd.DataFrame | None:
    try:
        payload = json.loads(text)
    except Exception:
        return None
    rows = payload.get("data") if isinstance(payload, dict) else None
    if rows is None:
        return None
    if not rows:
        return pd.DataFrame(columns=["mlbam"])
    df = pd.DataFrame(rows)
    idcol = next((c for c in df.columns if c.lower() == "xmlbamid"), None)
    if idcol is None:
        return None
    df = df.rename(columns={idcol: "mlbam"})
    df["mlbam"] = pd.to_numeric(df["mlbam"], errors="coerce")
    df = df.dropna(subset=["mlbam"])
    df["mlbam"] = df["mlbam"].astype(int)
    # drop HTML-anchor columns (Name/Team) and name/season duplicates
    html_cols = [c for c in df.columns if df[c].astype(str).str.contains("<a ", na=False).any()]
    drop = set(html_cols) | {"PlayerName", "PlayerNameRoute", "Name", "Team", "TeamNameAbb", "teamid", "Season", "SeasonMin", "SeasonMax"}
    df = df.drop(columns=[c for c in df.columns if c in drop])
    df = df.drop_duplicates("mlbam")
    return df.rename(columns={c: f"FG_{c}" for c in df.columns if c != "mlbam"})


def fg_dataset_year(stats: str, type_id: int, year: int) -> pd.DataFrame | None:
    cname = f"fg_{stats}_{type_id}_{year}.pkl"
    cached = load_cached(cname, year)
    if cached is not None:
        return cached
    text = http_text(_fg_url(stats, year, type_id), delay=FANGRAPHS_DELAY,
                     headers={"User-Agent": BROWSER_UA, "Accept": "application/json"})
    if text is None:
        return None
    df = parse_fangraphs(text)
    if df is not None:
        save_cached(cname, df)
    return df


def load_fangraphs() -> dict[str, pd.DataFrame]:
    print(f"\n[FanGraphs] Best-effort pull {FANGRAPHS_START_YEAR}-{CURRENT_YEAR} "
          f"({len(FANGRAPHS_TYPES)} tables x 2 groups per year; ~{FANGRAPHS_DELAY:.0f}s between requests)...")
    out: dict[str, list] = {"hit": [], "pit": []}
    ok = fail = 0
    for year in range(FANGRAPHS_START_YEAR, CURRENT_YEAR + 1):
        counts = []
        for group, stats in (("hit", "bat"), ("pit", "pit")):
            wide = None
            for type_id in FANGRAPHS_TYPES:
                df = fg_dataset_year(stats, type_id, year)
                if df is None:
                    fail += 1
                    if ok == 0 and fail >= 3:
                        print("[FanGraphs WARNING] Requests are being blocked/failing. Skipping FanGraphs.")
                        return {"hit": pd.DataFrame(), "pit": pd.DataFrame()}
                    continue
                ok += 1
                if df.empty:
                    continue
                wide = df if wide is None else merge_new_cols(wide, df)
            if wide is not None and not wide.empty:
                wide = wide.copy()
                wide["Year"] = year
                out[group].append(wide)
                counts.append(f"{group}: {len(wide)} players, {wide.shape[1] - 2} cols")
        print(f"-> FanGraphs {year}: " + (" | ".join(counts) if counts else "no data"))
    return {g: (pd.concat(f, ignore_index=True) if f else pd.DataFrame()) for g, f in out.items()}


# =============================== Source: draft + Chadwick IDs ===============================

def load_draft() -> dict[int, pd.DataFrame]:
    """Every draft pick (school, round, team) from the MLB API, grouped by player."""
    cached = load_cached("draft_all.pkl")
    if cached is not None:
        print(f"[Draft] Using cached draft table ({len(cached)} picks)")
        return index_by_player(cached)
    print(f"\n[Draft] Pulling MLB drafts {DRAFT_START_YEAR}-{CURRENT_YEAR}...")
    rows = []
    for y in range(DRAFT_START_YEAR, CURRENT_YEAR + 1):
        data = api_get(f"/draft/{y}")
        if not data:
            continue
        drafts = data.get("drafts") or {}
        for rnd in drafts.get("rounds", []):
            for pick in rnd.get("picks", []):
                pid = (pick.get("person") or {}).get("id")
                if not pid:
                    continue
                school = pick.get("school") or {}
                rows.append({
                    "mlbam": pid,
                    "Draft_Year": pick.get("year") or drafts.get("draftYear") or y,
                    "Draft_Round": pick.get("pickRound"),
                    "Draft_Pick_Overall": pick.get("pickNumber"),
                    "Draft_Team": (pick.get("team") or {}).get("name"),
                    "Draft_School": school.get("name"),
                    "Draft_School_Class": school.get("schoolClass"),
                    "Draft_School_State": school.get("state"),
                    "Draft_School_Country": school.get("country"),
                    "Draft_Signing_Bonus": pick.get("signingBonus"),
                })
    if not rows:
        print("[Draft WARNING] No draft data retrieved.")
        return {}
    df = pd.DataFrame(rows)
    save_cached("draft_all.pkl", df)
    print(f"[Draft] {len(df)} picks")
    return index_by_player(df)


def load_chadwick() -> dict[int, dict]:
    """ID cross-reference (BBRef, FanGraphs, Retrosheet) keyed by MLBAM id."""
    cached = load_cached("chadwick_ids.pkl")
    if cached is None:
        print("\n[Chadwick] Downloading player ID register (16 files, one-time)...")
        parts = []
        for h in "0123456789abcdef":
            text = http_text(CHADWICK_URL.format(h), delay=0.3, timeout=180)
            if text is None:
                print(f"[Chadwick WARNING] people-{h}.csv failed; skipping ID cross-reference.")
                return {}
            df = pd.read_csv(io.StringIO(text), low_memory=False)
            keep = [c for c in ("key_mlbam", "key_bbref", "key_bbref_minors", "key_fangraphs", "key_retro") if c in df.columns]
            parts.append(df[keep].dropna(subset=["key_mlbam"]))
        cached = pd.concat(parts, ignore_index=True)
        cached["key_mlbam"] = cached["key_mlbam"].astype(int)
        cached = cached.drop_duplicates("key_mlbam")
        save_cached("chadwick_ids.pkl", cached)
    return {int(r["key_mlbam"]): {k: v for k, v in r.items() if k != "key_mlbam"} for r in cached.to_dict("records")}


# =============================== MLB API: stats, awards, bio ===============================

HYDRATE_CONFIGS = [(True, True), (False, True), (True, False), (False, False)]   # (advanced, fielding)


def _hydrate(sport_id: int) -> str:
    groups = "hitting,pitching,fielding" if USE_FIELDING else "hitting,pitching"
    types = "yearByYear,yearByYearAdvanced" if USE_ADVANCED else "yearByYear"
    return f"stats(group=[{groups}],type=[{types}],sportId={sport_id})"


def fetch_batch(ids: list[int], sport_id: int) -> dict[int, dict]:
    """Bio + year-by-year stats for a batch of players at one level. Downgrades the hydrate on 4xx."""
    global USE_ADVANCED, USE_FIELDING, LAST_STATUS
    for _ in range(len(HYDRATE_CONFIGS)):
        LAST_STATUS = None
        data = api_get("/people", {"personIds": ",".join(map(str, ids)), "hydrate": _hydrate(sport_id)})
        if data is not None:
            return {p["id"]: p for p in data.get("people", [])}
        if LAST_STATUS is None or not (400 <= LAST_STATUS < 500):
            break   # network trouble, not a bad hydrate - don't downgrade
        if USE_ADVANCED:
            print("[INFO] Hydrate rejected; retrying without yearByYearAdvanced.")
            USE_ADVANCED = False
        elif USE_FIELDING:
            print("[INFO] Hydrate rejected again; retrying without fielding.")
            USE_FIELDING = False
        else:
            break
    return {}


def fetch_awards(ids: list[int]) -> dict[int, list[dict]]:
    data = api_get("/people", {"personIds": ",".join(map(str, ids)), "hydrate": "awards"})
    return {p["id"]: p.get("awards", []) for p in (data or {}).get("people", [])}


def extract_rows(person: dict, level: str) -> list[dict]:
    """Flatten one person's stat blocks at one level into row dicts (one per season/team[/position])."""
    rows: dict[tuple, dict] = {}
    for block in person.get("stats", []):
        group = block.get("group", {}).get("displayName")
        if group not in ("hitting", "pitching", "fielding"):
            continue
        seen: dict[tuple, int] = {}
        for s in block.get("splits", []):
            try:
                year = int(s.get("season"))
            except (TypeError, ValueError):
                continue
            stat = s.get("stat") or {}
            team = s.get("team")
            team_id = team.get("id") if team else None
            pos = (stat.get("position") or {}).get("abbreviation") if group == "fielding" else None
            occ_key = (year, team_id, pos)
            occ = seen.get(occ_key, 0)
            seen[occ_key] = occ + 1
            row = rows.setdefault((group, year, level, team_id, pos, occ), {
                "Group": group,
                "Year": year,
                "Level": level,
                "Team": team.get("name") if team else "TOT",
                "Team_ID": team_id,
                "League": (s.get("league") or {}).get("name"),
                "Is_Total_Row": team is None,
                "Is_Synthetic_Total": False,
                "Num_Teams": s.get("numTeams"),
                **({"Position": pos} if group == "fielding" else {}),
            })
            for k, v in stat.items():
                if isinstance(v, (dict, list)):
                    continue
                row.setdefault(k, v)
    return list(rows.values())


def bio_row(person: dict, chad: dict | None, draft: pd.DataFrame | None) -> pd.DataFrame:
    row: dict = {"MLBAM_ID": person.get("id"), "Player": person.get("fullName")}
    for k, v in person.items():
        if k in ("stats", "awards", "id", "fullName", "link", "copyright"):
            continue
        if v is None or isinstance(v, (str, int, float, bool)):
            row[k] = v
        elif isinstance(v, dict):
            for kk in ("abbreviation", "description", "name", "code"):
                if kk in v:
                    row[f"{k}_{kk}"] = v[kk]
    for k, v in (chad or {}).items():
        row[k] = v
    if draft is not None and not draft.empty:
        last = draft.sort_values("Draft_Year").iloc[-1]
        for k, v in last.items():
            row[k] = v
    return pd.DataFrame([row])


def awards_frames(awards: list[dict] | None) -> tuple[pd.DataFrame | None, pd.DataFrame | None]:
    """(raw awards rows for the Awards sheet, per-year summary to merge onto MLB rows)."""
    if not awards:
        return None, None
    rows = []
    for a in awards:
        try:
            year = int(a.get("season"))
        except (TypeError, ValueError):
            year = None
        rows.append({
            "Year": year,
            "Award_ID": a.get("id"),
            "Award": a.get("name"),
            "Date": a.get("date"),
            "Team": (a.get("team") or {}).get("name"),
            "Notes": a.get("notes"),
        })
    raw = pd.DataFrame(rows).sort_values("Year", na_position="last").reset_index(drop=True)
    dated = raw.dropna(subset=["Year"]).copy()
    if dated.empty:
        return raw, None
    dated["Year"] = dated["Year"].astype(int)
    summ = dated.groupby("Year").agg(
        Awards=("Award", lambda s: "; ".join(dict.fromkeys(x for x in s if isinstance(x, str)))),
        All_Star=("Award", lambda s: any(isinstance(x, str) and "all-star" in x.lower() for x in s)),
    ).reset_index()
    return raw, summ


# =============================== Frame building ===============================

def _ip_to_outs(ip) -> int:
    """'180.2' means 180 and 2/3 innings -> outs."""
    try:
        whole, _, frac = str(ip or "0").partition(".")
        return int(whole) * 3 + int(frac or 0)
    except ValueError:
        return 0


def age_on_april_1(birth: str | None, year: int) -> tuple[int | None, float | None]:
    if not birth:
        return None, None
    try:
        b = datetime.strptime(birth, "%Y-%m-%d").date()
    except ValueError:
        return None, None
    whole = year - b.year - (1 if (b.month, b.day) > (4, 1) else 0)
    return whole, round((date(year, 4, 1) - b).days / 365.25, 2)


def _sort_rows(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_lvl"] = df["Level"].map(LEVEL_RANK)
    df["_tot"] = (~df["Is_Total_Row"].astype(bool)).astype(int)
    cols = ["Year", "_lvl", "_tot", "Team"] + (["Position"] if "Position" in df.columns else [])
    df = df.sort_values(cols, kind="stable").drop(columns=["_lvl", "_tot"])
    return df.reset_index(drop=True)


def build_frame(rows: list[dict], group: str, person: dict) -> pd.DataFrame:
    rs = [r for r in rows if r["Group"] == group]
    if not rs:
        return pd.DataFrame()
    df = pd.DataFrame(rs).drop(columns=["Group"])
    stat_cols = [c for c in df.columns if c not in META_COLS]

    if group == "pitching" and "inningsPitched" in df.columns:
        df["IP"] = (df["inningsPitched"].apply(_ip_to_outs) / 3).round(3)

    for c in stat_cols:
        if c in ("inningsPitched", "innings"):
            continue
        col = df[c].replace(PLACEHOLDERS, pd.NA)
        conv = pd.to_numeric(col, errors="coerce")
        if conv.notna().sum() >= col.notna().sum() * 0.9:
            df[c] = conv
    df = df.copy()   # consolidate blocks (avoids PerformanceWarning + speeds up later column adds)

    if group == "hitting":
        if {"hits", "doubles", "triples", "homeRuns"} <= set(df.columns):
            df["singles"] = df["hits"] - df["doubles"] - df["triples"] - df["homeRuns"]
        if {"stolenBases", "caughtStealing"} <= set(df.columns):
            df["stolenBaseAttempts"] = df["stolenBases"] + df["caughtStealing"]

    df["Player"] = person.get("fullName")
    df["MLBAM_ID"] = person.get("id")
    ages = df["Year"].apply(lambda y: age_on_april_1(person.get("birthDate"), y))
    df["Age_Apr1"] = ages.apply(lambda a: a[0])
    df["Age_Apr1_Decimal"] = ages.apply(lambda a: a[1])

    df = _sort_rows(df)
    ordered = [c for c in META_COLS if c in df.columns] + [c for c in df.columns if c not in META_COLS]
    return df[ordered]


def attach_all(df: pd.DataFrame, exts: list[pd.DataFrame | None]) -> pd.DataFrame:
    """
    Merge per-year external frames (each: 'Year' + data columns, one row per year) onto the MLB rows.
    Target row per year: the combined (TOT) row if the API gave one, else the only MLB row. For a
    multi-team year with no combined row, a TOT row is created (Is_Synthetic_Total=True) to hold the
    external stats; its API stat columns are blank.
    """
    exts = [e for e in exts if e is not None and not e.empty and "Year" in e.columns]
    if df.empty or not exts:
        return df
    df = df.copy()
    years = sorted({int(y) for e in exts for y in e["Year"].dropna()})

    synth = []
    for y in years:
        mlb = df[(df["Year"] == y) & (df["Level"] == "MLB")]
        if len(mlb) > 1 and not mlb["Is_Total_Row"].astype(bool).any():
            base = mlb.iloc[0]
            rec = {c: base[c] for c in ("Player", "MLBAM_ID", "Year", "Age_Apr1", "Age_Apr1_Decimal", "Level") if c in df.columns}
            rec.update({"Team": "TOT", "Team_ID": None, "League": None, "Is_Total_Row": True,
                        "Is_Synthetic_Total": True, "Num_Teams": len(mlb)})
            synth.append(rec)
    if synth:
        df = _sort_rows(pd.concat([df, pd.DataFrame(synth)], ignore_index=True))

    targets: dict[int, int] = {}
    for y in years:
        mlb = df[(df["Year"] == y) & (df["Level"] == "MLB")]
        if mlb.empty:
            continue
        tot = mlb[mlb["Is_Total_Row"].astype(bool)]
        if not tot.empty:
            targets[y] = tot.index[0]
        elif len(mlb) == 1:
            targets[y] = mlb.index[0]

    for e in exts:
        e2 = e.copy()
        e2["_idx"] = e2["Year"].map(lambda y: targets.get(int(y)) if pd.notna(y) else None)
        e2 = e2.dropna(subset=["_idx"])
        if e2.empty:
            continue
        e2["_idx"] = e2["_idx"].astype(int)
        e2 = e2.drop(columns=["Year"]).drop_duplicates("_idx").set_index("_idx")
        new = [c for c in e2.columns if c not in df.columns]
        if new:
            df = df.join(e2[new])
    return df


# =============================== Classification + saving ===============================

def _canonical_mlb(df: pd.DataFrame) -> pd.DataFrame:
    """MLB rows with traded-player double counting removed (prefer the API's combined row per year)."""
    if df.empty:
        return df
    mlb = df[df["Level"] == "MLB"]
    keep = []
    for _, g in mlb.groupby("Year"):
        tot = g[g["Is_Total_Row"].astype(bool)]
        keep.append(tot if not tot.empty else g)
    return pd.concat(keep) if keep else mlb


def is_hitter(hit: pd.DataFrame, pit: pd.DataFrame) -> bool:
    """MLB plate appearances vs innings pitched; two-way players land where they have more volume."""
    h, p = _canonical_mlb(hit), _canonical_mlb(pit)
    pa = h["plateAppearances"].sum() if "plateAppearances" in h.columns else 0
    ip = p["IP"].sum() if "IP" in p.columns else 0
    return bool(pa > ip)


def _safe_name(name: str) -> str:
    return re.sub(r"[^\w\-]+", "_", str(name)).strip("_")


def save_player(person: dict, rows: list[dict], ctx: dict, awards: list[dict] | None) -> bool:
    pid = person["id"]
    hit = build_frame(rows, "hitting", person)
    pit = build_frame(rows, "pitching", person)
    fld = build_frame(rows, "fielding", person)
    if hit.empty and pit.empty and fld.empty:
        return False

    hitter = is_hitter(hit, pit)          # classify on API rows only, before external merges

    bat_raw, pit_raw = ctx["bref_bat"].get(pid), ctx["bref_pit"].get(pid)
    awards_raw, awards_year = awards_frames(awards)

    hit = attach_all(hit, [aggregate_bref(bat_raw, "OPS_plus", "PA"), ctx["sv_hit"].get(pid),
                           ctx["fg_hit"].get(pid), awards_year])
    pit = attach_all(pit, [aggregate_bref(pit_raw, "ERA_plus", "IPouts"), ctx["sv_pit"].get(pid),
                           ctx["fg_pit"].get(pid), awards_year])

    draft = ctx["draft"].get(pid)
    bio = bio_row(person, ctx["chadwick"].get(pid), draft)

    out_dir, other_dir = (HITTERS_DIR, PITCHERS_DIR) if hitter else (PITCHERS_DIR, HITTERS_DIR)
    filename = f"{pid}_{_safe_name(person.get('fullName'))}.xlsx"
    stale = os.path.join(other_dir, filename)
    if os.path.exists(stale):
        os.remove(stale)

    sheets = [("Bio", bio), ("Hitting", hit), ("Pitching", pit), ("Fielding", fld), ("Awards", awards_raw),
              ("Draft", draft), ("BRef_Hitting_WAR", bat_raw), ("BRef_Pitching_WAR", pit_raw)]
    with pd.ExcelWriter(os.path.join(out_dir, filename), engine=EXCEL_ENGINE) as xw:
        for name, frame in sheets:
            if frame is not None and not frame.empty:
                frame.to_excel(xw, sheet_name=name, index=False)
    return True


def process_batch(batch: list[int], ctx: dict) -> tuple[int, int]:
    by_sport = {sport: fetch_batch(batch, sport) for sport in LEVELS}
    awards_by_id = fetch_awards(batch) if SOURCES["awards"] else {}
    saved = failed = 0
    for pid in batch:
        try:
            rows: list[dict] = []
            person = None
            for sport, level in LEVELS.items():
                p = by_sport[sport].get(pid)
                if p is None:
                    continue
                person = person or p
                rows += extract_rows(p, level)
            if person is None:
                failed += 1
                continue
            if save_player(person, rows, ctx, awards_by_id.get(pid)):
                saved += 1
            else:
                failed += 1
        except Exception as exc:
            failed += 1
            print(f"[ERROR] player {pid}: {exc}")
    return saved, failed


def main():
    setup_directories()
    empty = pd.DataFrame()

    bref_bat = load_bref(BREF_BAT_URL, "war_daily_bat.csv") if SOURCES["bref"] else {}
    bref_pit = load_bref(BREF_PIT_URL, "war_daily_pitch.csv") if SOURCES["bref"] else {}
    savant = load_savant() if SOURCES["savant"] else {"hit": empty, "pit": empty}
    fangraphs = load_fangraphs() if SOURCES["fangraphs"] else {"hit": empty, "pit": empty}
    draft = load_draft() if SOURCES["draft"] else {}
    chadwick = load_chadwick() if SOURCES["chadwick"] else {}

    ctx = {
        "bref_bat": bref_bat, "bref_pit": bref_pit,
        "sv_hit": index_by_player(savant["hit"]), "sv_pit": index_by_player(savant["pit"]),
        "fg_hit": index_by_player(fangraphs["hit"]), "fg_pit": index_by_player(fangraphs["pit"]),
        "draft": draft, "chadwick": chadwick,
    }

    print("\n[Sources loaded] " + ", ".join([
        f"BRef bat {len(bref_bat)} / pit {len(bref_pit)} players",
        f"Savant hit {len(ctx['sv_hit'])} / pit {len(ctx['sv_pit'])}",
        f"FanGraphs hit {len(ctx['fg_hit'])} / pit {len(ctx['fg_pit'])}",
        f"draft {len(draft)}", f"Chadwick IDs {len(chadwick)}"]))

    player_ids = discover_player_ids(MLB_START_YEAR, CURRENT_YEAR)
    if not player_ids:
        print("Exiting: no players found. Check the warnings above.")
        return

    print(f"\n===================================================\nFetching AA/AAA/MLB stats for {len(player_ids)} players...\n===================================================")
    total_saved = total_failed = 0
    for i in range(0, len(player_ids), BATCH_SIZE):
        batch = player_ids[i:i + BATCH_SIZE]
        saved, failed = process_batch(batch, ctx)
        total_saved += saved
        total_failed += failed
        print(f"-> {min(i + BATCH_SIZE, len(player_ids))}/{len(player_ids)} processed "
              f"(saved {total_saved}, failed {total_failed}) | advanced: {USE_ADVANCED}, fielding: {USE_FIELDING}")

    print(f"\nDone. Saved {total_saved}, failed {total_failed}.")


if __name__ == "__main__":
    main()