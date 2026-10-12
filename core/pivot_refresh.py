from __future__ import annotations

from pathlib import Path
import re
from typing import Any, cast


class PivotRefreshError(RuntimeError):
    pass


def _add_missing_pivot(
    pivot_cache: Any,
    worksheet: Any,
    table_name: str,
    is_summary: bool,
) -> None:
    pivot = pivot_cache.CreatePivotTable(
        TableDestination=worksheet.Range("A3"),
        TableName=table_name,
    )
    room_field = pivot.PivotFields("Specific Location (Room No.)")
    room_field.Orientation = 1
    room_field.Position = 1

    if is_summary:
        description_field = pivot.PivotFields("Asset Description")
        description_field.Orientation = 1
        description_field.Position = 2

    pivot.AddDataField(
        pivot.PivotFields("Asset Description"),
        "Asset Count",
        -4112,
    )


def _find_worksheet(workbook: Any, wanted_name: str) -> Any:
    key = re.sub(r"[^a-z0-9]", "", wanted_name.casefold())
    for index in range(1, workbook.Worksheets.Count + 1):
        worksheet = workbook.Worksheets(index)
        actual = re.sub(r"[^a-z0-9]", "", str(worksheet.Name).casefold())
        if actual == key:
            return worksheet
    raise PivotRefreshError(f"Required PivotTable worksheet not found: {wanted_name}")


def refresh_pivots(workbook_path: str | Path) -> None:
    path = Path(workbook_path).resolve()
    if not path.is_file():
        raise PivotRefreshError(f"Workbook to refresh does not exist: {path}")

    try:
        import win32com.client
    except ImportError as exc:
        raise PivotRefreshError(
            "Pivot refresh requires pywin32 and desktop Microsoft Excel on Windows. "
            "Install pywin32 in the application's Python environment and retry."
        ) from exc

    excel: Any | None = None
    workbook: Any | None = None
    try:
        excel = cast(Any, win32com.client.DispatchEx("Excel.Application"))
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = cast(
            Any,
            excel.Workbooks.Open(
                str(path),
                UpdateLinks=0,
                ReadOnly=False,
            ),
        )

        listing: Any = _find_worksheet(workbook, "Listing - File")
        last_column = max(
            1,
            cast(int, listing.Cells(1, listing.Columns.Count).End(-4159).Column),
        )
        last_row = max(
            cast(int, listing.Cells(listing.Rows.Count, column).End(-4162).Row)
            for column in range(1, last_column + 1)
        )
        source_data = (
            f"'Listing - File'!R1C1:R{last_row}C{last_column}"
        )

        caches: Any = cast(Any, workbook.PivotCaches())
        pivot_cache: Any
        if caches.Count:
            pivot_cache = cast(Any, caches.Item(1))
            for index in range(1, caches.Count + 1):
                cache: Any = cast(Any, caches.Item(index))
                cache.SourceData = source_data
        else:
            pivot_cache = cast(
                Any,
                caches.Create(
                    SourceType=1,
                    SourceData=source_data,
                    Version=6,
                ),
            )

        required_sheets = (
            ("Summary-Listing", True),
            ("Location Count - Per File", False),
        )
        for sheet_name, is_summary in required_sheets:
            worksheet = _find_worksheet(workbook, sheet_name)
            if worksheet.PivotTables().Count == 0:
                _add_missing_pivot(
                    pivot_cache,
                    worksheet,
                    f"Inventory_{'Summary' if is_summary else 'LocationCount'}",
                    is_summary,
                )

        workbook.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        for sheet_name in ("Summary-Listing", "Location Count - Per File"):
            worksheet = _find_worksheet(workbook, sheet_name)
            if worksheet.PivotTables().Count == 0:
                raise PivotRefreshError(
                    f"Excel did not create or retain a PivotTable on '{worksheet.Name}'."
                )
        workbook.Save()
    except Exception as exc:
        raise PivotRefreshError(
            f"Microsoft Excel could not refresh and save PivotTables in '{path}': {exc}"
        ) from exc
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        finally:
            if excel is not None:
                excel.Quit()
