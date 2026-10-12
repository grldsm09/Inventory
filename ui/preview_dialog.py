from __future__ import annotations

from collections.abc import Iterable
from typing import Any

import pandas as pd
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QHeaderView,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)


class PreviewDialog(QDialog):
    def __init__(
        self,
        title: str,
        values: pd.DataFrame | Iterable[object],
        parent: Any = None,
        max_rows: int = 500,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(1000, 620)

        layout = QVBoxLayout(self)
        if isinstance(values, pd.DataFrame):
            preview = values.head(max_rows)
            table = QTableWidget(len(preview), len(preview.columns), self)
            table.setHorizontalHeaderLabels([str(column) for column in preview.columns])
            for row_index, row in enumerate(preview.itertuples(index=False, name=None)):
                for column_index, value in enumerate(row):
                    table.setItem(
                        row_index,
                        column_index,
                        QTableWidgetItem("" if pd.isna(value) else str(value)),
                    )
            total = len(values)
            shown = len(preview)
        else:
            items = list(values)
            table = QTableWidget(len(items), 1, self)
            table.setHorizontalHeaderLabels(["Location"])
            for row_index, value in enumerate(items):
                table.setItem(row_index, 0, QTableWidgetItem(str(value)))
            total = len(items)
            shown = len(items)

        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(QLabel(f"Showing {shown:,} of {total:,} records"))
        layout.addWidget(table)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)
