from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from typing import Iterable, Mapping

import pandas as pd


def normalize_location(value: object) -> str:
    """Normalize case, Unicode composition, and whitespace without merging punctuation."""
    if value is None or pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    return " ".join(text.split()).casefold()


def normalize_header(value: object) -> str:
    text = unicodedata.normalize("NFKC", str(value)).casefold()
    return "".join(character for character in text if character.isalnum())


def display_location(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return " ".join(unicodedata.normalize("NFKC", str(value)).split())


def normalize_worksheet_key(value: object) -> str:
    text = unicodedata.normalize("NFKC", str(value)).casefold()
    return "".join(character for character in text if character.isalnum())


FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "main_location": (
        "Main Location (Bldg.)",
        "Main Location (Building)",
        "Main Location",
    ),
    "specific_location": (
        "Specific Location (Room No.)",
        "Specific Location",
        "Room No.",
        "Room",
    ),
    "asset_tag": ("Asset Tag No", "Asset Tag", "Tag No."),
    "asset_description": ("Asset Description", "Description"),
    "quantity": ("Qty", "Quantity"),
    "supplier": ("Supplier", "Vendor"),
    "acquisition_date": ("Acquisition Date",),
    "unit_of_measure": ("UOM", "Unit of Measure"),
    "account_charge": ("Account Charge",),
    "asset_category": ("Asset Category",),
    "type_of_asset": ("Type of Asset",),
    "classification": ("Classification",),
    "date_tagged": ("Date Tagged",),
    "unit_cost": ("Unit Cost",),
    "total_cost": ("Total Cost",),
    "rr_number": ("RR No", "RR Number"),
    "po_number": ("PO NO.", "PO No", "PO Number"),
    "serial_number": ("Serial No", "Serial Number"),
    "sub_location": ("Sub Location (Floor)", "Sub Location", "Floor"),
    "department": ("DEPT", "Department"),
    "rc_code": ("RC CODE", "RC Code"),
    "program": ("PROGRAM", "Program"),
    "pc_code": ("PC CODE", "PC Code"),
    "custodian": ("Custodian",),
    "remarks": ("Remarks",),
}


def resolve_column(columns: Iterable[object], field: str) -> object | None:
    """Resolve a known inventory field against real source or workbook headings."""
    column_list = list(columns)
    aliases = FIELD_ALIASES.get(field)
    if aliases is None:
        aliases = (field,)

    alias_keys = {normalize_header(alias) for alias in aliases}
    for column in column_list:
        if normalize_header(column) in alias_keys:
            return column

    if field == "main_location":
        for column in column_list:
            key = normalize_header(column)
            if "mainlocation" in key and ("bldg" in key or "building" in key or key == "mainlocation"):
                return column
    elif field == "specific_location":
        for column in column_list:
            key = normalize_header(column)
            if "specificlocation" in key or key in {"room", "roomno"}:
                return column
    return None


@dataclass(frozen=True)
class LocationSheet:
    key: str
    display_name: str
    worksheet_name: str


class LocationMapper:
    def __init__(
        self,
        values: Iterable[object],
        overrides: Mapping[str, str] | None = None,
    ) -> None:
        self._overrides = {
            normalize_location(key): display_location(value)
            for key, value in (overrides or {}).items()
            if normalize_location(key) and display_location(value)
        }
        self._alias_to_display: dict[str, str] = {}
        self._source_labels: dict[str, str] = {}
        self._source_variants: dict[str, list[str]] = {}

        for value in values:
            source_label = display_location(value)
            source_key = normalize_location(source_label)
            if not source_key:
                continue
            self._source_labels.setdefault(source_key, source_label)
            variants = self._source_variants.setdefault(source_key, [])
            if source_label not in variants:
                variants.append(source_label)
            requested = self._overrides.get(source_key, self._source_labels[source_key])
            self._alias_to_display[source_key] = requested

        self._locations: dict[str, str] = {}
        for source_key, target in self._alias_to_display.items():
            target_key = normalize_location(target)
            if target_key:
                self._locations.setdefault(target_key, target)

    @property
    def locations(self) -> list[str]:
        return list(self._locations.values())

    @property
    def location_keys(self) -> set[str]:
        return set(self._locations)

    def resolve(self, value: object) -> str:
        key = normalize_location(value)
        if not key:
            return ""
        return self._alias_to_display.get(key, display_location(value))

    def source_groups(self) -> list[tuple[str, list[str]]]:
        return [
            (
                self._overrides.get(key, self._source_labels[key]),
                variants,
            )
            for key, variants in self._source_variants.items()
        ]

    def worksheet_names(self, reserved_names: Iterable[str] = ()) -> dict[str, str]:
        reserved = {normalize_worksheet_key(name) for name in reserved_names}
        result: dict[str, str] = {}
        used = set(reserved)
        for key, display_name in self._locations.items():
            base = re.sub(r"[\[\]:*?/\\]", "-", display_name).strip().strip("'")
            base = base or "Location"
            base = base[:31]
            candidate = base
            suffix = 2
            while normalize_worksheet_key(candidate) in used:
                ending = f" ({suffix})"
                candidate = f"{base[:31 - len(ending)]}{ending}"
                suffix += 1
            used.add(normalize_worksheet_key(candidate))
            result[key] = candidate
        return result
