from pathlib import Path

APP_NAME = "Inventory Generator"
PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATE = (
    PROJECT_ROOT / "template" / "Inventory Location and Listings_EINSTEIN.xlsx"
)
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output"

SOURCE_SHEETS = (
    ("Asset Register 2025", 3),
    ("Asset Register 2026", 5),
)

SPECIAL_SHEET_KEYS = {
    "raw": "Raw",
    "listingfile": "Listing - File",
    "summarylisting": "Summary-Listing",
    "locationcountperfile": "Location Count - Per File",
}

ROOM_HEADER_ROW = 3
ROOM_DATA_ROW = ROOM_HEADER_ROW + 1
LISTING_HEADER_ROW = 1
LISTING_DATA_ROW = LISTING_HEADER_ROW + 1
RAW_HEADER_ROW = 1
RAW_DATA_ROW = RAW_HEADER_ROW + 1
MAX_PREVIEW_ROWS = 500
