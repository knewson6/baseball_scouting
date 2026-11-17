"""Create prospect-ready copies of the raw player workbooks."""

from pathlib import Path
import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
RAW_DATA_DIR = PROJECT_ROOT / "raw data"
OUTPUT_DIR = Path(__file__).resolve().parent / "prospect data"
HITTERS_PATH = OUTPUT_DIR / "hitters"
PITCHERS_PATH = OUTPUT_DIR / "pitchers"

LEVELS_TO_KEEP = {"AA", "AAA", "MLB"}
MLB_ONLY_PREFIXES = ("BRef_", "SV_")
MLB_ONLY_COLUMNS = {"Awards", "All_Star"}


def setup_directories() -> None:
    """Ensure the cleaned hitter and pitcher directories exist."""
    HITTERS_PATH.mkdir(parents=True, exist_ok=True)
    PITCHERS_PATH.mkdir(parents=True, exist_ok=True)


def clean_stats_sheet(data: pd.DataFrame) -> pd.DataFrame:
    """Keep selected levels and remove fields unavailable outside MLB."""
    if "Level" not in data.columns:
        raise ValueError("the stats sheet does not contain a Level column")

    levels = data["Level"].astype("string").str.upper()
    cleaned = data.loc[levels.isin(LEVELS_TO_KEEP)].copy()

    columns_to_drop = [
        column
        for column in cleaned.columns
        if column in MLB_ONLY_COLUMNS
        or column.startswith(MLB_ONLY_PREFIXES)
    ]
    return cleaned.drop(columns=columns_to_drop)


def clean_player_workbook(
    source_path: Path, destination_path: Path, stats_sheet: str
) -> None:
    """Write a player workbook containing Bio and cleaned hitting or pitching data."""
    with pd.ExcelFile(source_path) as workbook:
        if stats_sheet not in workbook.sheet_names:
            raise ValueError(f"missing required {stats_sheet!r} sheet")

        stats = pd.read_excel(workbook, sheet_name=stats_sheet)
        cleaned_stats = clean_stats_sheet(stats)
        bio = (
            pd.read_excel(workbook, sheet_name="Bio")
            if "Bio" in workbook.sheet_names
            else None
        )

    destination_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(destination_path, engine="xlsxwriter") as writer:
        if bio is not None:
            bio.to_excel(writer, sheet_name="Bio", index=False)
        cleaned_stats.to_excel(writer, sheet_name=stats_sheet, index=False)


def clean_player_directory(
    source_dir: Path,
    destination_dir: Path,
    stats_sheet: str,
    alternate_sheet: str,
    alternate_destination: Path,
) -> tuple[int, int, int]:
    """Clean every Excel workbook in one raw-data directory."""
    processed = 0
    skipped = 0
    rerouted = 0

    for source_path in sorted(source_dir.glob("*.xlsx")):
        try:
            with pd.ExcelFile(source_path) as workbook:
                available_sheets = set(workbook.sheet_names)

            selected_sheet = stats_sheet
            selected_destination = destination_dir
            if selected_sheet not in available_sheets:
                if alternate_sheet not in available_sheets:
                    raise ValueError(f"missing required {stats_sheet!r} sheet")
                selected_sheet = alternate_sheet
                selected_destination = alternate_destination
                rerouted += 1

            clean_player_workbook(
                source_path,
                selected_destination / source_path.name,
                selected_sheet,
            )
        except (OSError, ValueError, ImportError) as error:
            skipped += 1
            print(f"Skipping {source_path.name}: {error}")
            continue

        processed += 1

    return processed, skipped, rerouted


def main() -> None:
    setup_directories()

    results = (
        (
            "hitters",
            RAW_DATA_DIR / "hitters",
            HITTERS_PATH,
            "Hitting",
        ),
        (
            "pitchers",
            RAW_DATA_DIR / "pitchers",
            PITCHERS_PATH,
            "Pitching",
        ),
    )

    for label, source_dir, destination_dir, stats_sheet in results:
        if not source_dir.is_dir():
            raise FileNotFoundError(f"Raw data directory not found: {source_dir}")

        processed, skipped, rerouted = clean_player_directory(
            source_dir,
            destination_dir,
            stats_sheet,
            "Pitching" if stats_sheet == "Hitting" else "Hitting",
            PITCHERS_PATH if stats_sheet == "Hitting" else HITTERS_PATH,
        )
        print(
            f"{label.capitalize()}: processed {processed} workbook(s); "
            f"skipped {skipped}; rerouted {rerouted}."
        )


if __name__ == "__main__":
    main()
