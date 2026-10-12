from __future__ import annotations

from typing import Any

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QHeaderView,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from core.mapper import LocationMapper, normalize_location


class MappingDialog(QDialog):
    def __init__(self, values: list[object], parent: Any = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Specific Location Mapping")
        self.resize(760, 560)
        self._rows: list[tuple[str, str]] = []

        mapper = LocationMapper(values)
        layout = QVBoxLayout(self)
        layout.addWidget(
            QLabel(
                "Review the automatically normalized room names. Edit a worksheet name "
                "to merge aliases or choose a clearer name."
            )
        )
        self.table = QTableWidget(len(mapper.source_groups()), 3, self)
        self.table.setHorizontalHeaderLabels(
            ["Normalized key", "Source spelling(s)", "Worksheet name"]
        )
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        for row, (canonical, sources) in enumerate(mapper.source_groups()):
            key = normalize_location(sources[0])
            self._rows.append((key, canonical))
            self.table.setItem(row, 0, QTableWidgetItem(key))
            self.table.setItem(row, 1, QTableWidgetItem(", ".join(sources)))
            self.table.setItem(row, 2, QTableWidgetItem(canonical))
        layout.addWidget(self.table)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def overrides(self) -> dict[str, str]:
        result: dict[str, str] = {}
        for row, (key, original) in enumerate(self._rows):
            item = self.table.item(row, 2)
            value = item.text().strip() if item else ""
            result[key] = value or original
        return result
