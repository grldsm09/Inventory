from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import pandas as pd

from core.mapper import resolve_column
from core.settings import SOURCE_SHEETS

ProgressCallback = Callable[[int, str], None]


class DataLoadError(RuntimeError):
    pass


@dataclass
class LoadedAssets:
    dataframe: pd.DataFrame
    source_path: Path
    main_location_column: object
    specific_location_column: object


def load_asset_register(
    source_path: str | Path,
    progress_callback: ProgressCallback | None = None,
) -> LoadedAssets:
    path = Path(source_path).expanduser().resolve()
    if not path.is_file():
        raise DataLoadError(f"Source workbook does not exist: {path}")

    frames: list[pd.DataFrame] = []
    try:
        for index, (sheet_name, header_row) in enumerate(SOURCE_SHEETS):
            if progress_callback:
                progress_callback(index * 35, f"Reading {sheet_name}...")
            frame = pd.read_excel(
                path,
                sheet_name=sheet_name,
                header=header_row,
                dtype=object,
                engine="openpyxl",
            )
            frame.columns = [str(column).strip() for column in frame.columns]
            frames.append(frame)
    except Exception as exc:
        raise DataLoadError(
            f"Could not read the required Asset Register sheets from '{path}': {exc}"
        ) from exc

    dataframe = pd.concat(frames, ignore_index=True, sort=False)
    dataframe = dataframe.dropna(how="all").reset_index(drop=True)
    main_column = resolve_column(dataframe.columns, "main_location")
    specific_column = resolve_column(dataframe.columns, "specific_location")
    missing = []
    if main_column is None:
        missing.append("Main Location (Building)")
    if specific_column is None:
        missing.append("Specific Location (Room No.)")
    if missing:
        raise DataLoadError(
            "The merged Asset Register is missing required column(s): "
            + ", ".join(missing)
        )

    if progress_callback:
        progress_callback(100, f"Loaded {len(dataframe):,} asset records")
    return LoadedAssets(
        dataframe=dataframe,
        source_path=path,
        main_location_column=main_column,
        specific_location_column=specific_column,
    )
