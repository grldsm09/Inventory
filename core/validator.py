from __future__ import annotations

from collections.abc import Iterable
from typing import Any

import pandas as pd

from core.mapper import (
    LocationMapper,
    display_location,
    normalize_location,
    resolve_column,
)


class ValidationError(ValueError):
    def __init__(self, message: str, missing_locations: Iterable[str] = ()) -> None:
        super().__init__(message)
        self.missing_locations = list(missing_locations)


def list_main_locations(dataframe: pd.DataFrame, column: object) -> list[str]:
    unique: dict[str, str] = {}
    for value in dataframe[column].tolist():
        label = display_location(value)
        key = normalize_location(label)
        if key:
            unique.setdefault(key, label)
    return sorted(unique.values(), key=str.casefold)


def filter_by_main_location(
    dataframe: pd.DataFrame,
    column: object,
    selected_location: str,
) -> pd.DataFrame:
    key = normalize_location(selected_location)
    mask = dataframe[column].map(normalize_location).eq(key)
    return dataframe.loc[mask].copy()


def get_specific_location_column(dataframe: pd.DataFrame) -> object:
    column = resolve_column(dataframe.columns, "specific_location")
    if column is None:
        raise ValidationError("Specific Location (Room No.) column was not found.")
    return column


def validate_location_coverage(
    expected_keys: set[str],
    created_sheet_by_key: dict[str, str],
) -> list[str]:
    return sorted(expected_keys - set(created_sheet_by_key), key=str.casefold)


def duplicate_assets(dataframe: pd.DataFrame) -> pd.DataFrame:
    tag_column = resolve_column(dataframe.columns, "asset_tag")
    if tag_column is None or dataframe.empty:
        return pd.DataFrame(columns=[*dataframe.columns, "Duplicate Reason"])

    tags = dataframe[tag_column].map(display_location)
    ignored = {"", "no tag", "unreadable tag", "n/a", "na", "-"}
    eligible = tags.map(normalize_location).map(lambda value: value not in ignored)
    duplicate_mask = eligible & tags.map(normalize_location).duplicated(keep=False)
    duplicates = dataframe.loc[duplicate_mask].copy()
    if not duplicates.empty:
        duplicates["Duplicate Reason"] = (
            "Repeated Asset Tag No.: "
            + duplicates[tag_column].map(display_location)
        )
    else:
        duplicates["Duplicate Reason"] = pd.Series(dtype=object)
    return duplicates


def validate_filtered_records(
    dataframe: pd.DataFrame,
    location_column: object,
    room_column: object,
    selected_location: str,
    mapper: LocationMapper,
) -> list[str]:
    if dataframe.empty:
        raise ValidationError(
            f"No asset records were found for Main Location '{selected_location}'."
        )

    actual_locations = {
        normalize_location(value)
        for value in dataframe[location_column].tolist()
        if normalize_location(value)
    }
    expected_location = normalize_location(selected_location)
    if actual_locations != {expected_location}:
        raise ValidationError(
            "Filtered records contain an unexpected Main Location. "
            "Generation stopped to prevent cross-building data."
        )

    blank_rows = dataframe[room_column].map(normalize_location).eq("")
    missing = ["<blank Specific Location>"] if blank_rows.any() else []
    if missing:
        raise ValidationError(
            f"{int(blank_rows.sum())} selected asset record(s) have no Specific Location. "
            "Assign a location before generating the workbook.",
            missing_locations=missing,
        )

    resolved_keys = {
        normalize_location(mapper.resolve(value))
        for value in dataframe[room_column].tolist()
        if mapper.resolve(value)
    }
    missing_keys = resolved_keys - mapper.location_keys
    return sorted(missing_keys, key=str.casefold)


def report_rows_for_missing(
    dataframe: pd.DataFrame,
    room_column: object,
    selected_location: str,
    missing_locations: Iterable[str],
) -> list[dict[str, Any]]:
    missing_set = {normalize_location(value) for value in missing_locations}
    rows: list[dict[str, Any]] = []
    if "<blank specific location>" in missing_set:
        missing_rows = dataframe.loc[dataframe[room_column].map(normalize_location).eq("")]
        for index, record in missing_rows.iterrows():
            rows.append(
                {
                    "Main Location": selected_location,
                    "Specific Location": "",
                    "Source Row": int(index) + 1,
                    "Issue": "Specific Location is blank",
                }
            )
    for value in missing_locations:
        if value != "<blank Specific Location>":
            rows.append(
                {
                    "Main Location": selected_location,
                    "Specific Location": value,
                    "Source Row": "",
                    "Issue": "Worksheet was not created",
                }
            )
    return rows
