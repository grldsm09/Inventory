from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from PySide6.QtCore import QSettings, Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QComboBox,
    QCompleter,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.loader import LoadedAssets, load_asset_register
from core.mapper import LocationMapper
from core.reports import export_processing_log, make_log_record
from core.settings import DEFAULT_OUTPUT_DIR, DEFAULT_TEMPLATE, MAX_PREVIEW_ROWS
from core.validator import filter_by_main_location, list_main_locations
from core.worker_thread import OperationWorker
from core.workbook_generator import GenerationResult, generate_workbook
from ui.mapping_dialog import MappingDialog
from ui.preview_dialog import PreviewDialog


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Inventory Generator")
        self.resize(1120, 800)
        self.settings = QSettings()
        self.loaded_assets: LoadedAssets | None = None
        self.location_overrides: dict[str, str] = {}
        self.log_records: list[dict[str, str]] = []
        self.worker: OperationWorker | None = None
        self.last_result: GenerationResult | None = None

        central = QWidget(self)
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        file_group = QGroupBox("Workbooks")
        file_layout = QGridLayout(file_group)
        self.source_path = QLabel("No Asset Monitoring workbook selected")
        self.template_path = QLabel(str(DEFAULT_TEMPLATE))
        self.output_path = QLabel(str(DEFAULT_OUTPUT_DIR))
        self.btn_browse_source = QPushButton("Browse Source")
        self.btn_browse_template = QPushButton("Browse Template")
        self.btn_browse_output = QPushButton("Output Folder")
        file_layout.addWidget(QLabel("Source"), 0, 0)
        file_layout.addWidget(self.source_path, 0, 1)
        file_layout.addWidget(self.btn_browse_source, 0, 2)
        file_layout.addWidget(QLabel("Template"), 1, 0)
        file_layout.addWidget(self.template_path, 1, 1)
        file_layout.addWidget(self.btn_browse_template, 1, 2)
        file_layout.addWidget(QLabel("Output"), 2, 0)
        file_layout.addWidget(self.output_path, 2, 1)
        file_layout.addWidget(self.btn_browse_output, 2, 2)
        root.addWidget(file_group)

        selection = QGroupBox("Building and Location")
        selection_layout = QGridLayout(selection)
        self.cmb_building = QComboBox()
        self.cmb_building.setEditable(True)
        self.cmb_building.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        completer = QCompleter(self.cmb_building.model(), self)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.cmb_building.setCompleter(completer)
        self.btn_preview_assets = QPushButton("Preview Assets")
        self.btn_preview_locations = QPushButton("Preview Locations")
        self.btn_mapping = QPushButton("Location Mapping")
        self.btn_generate = QPushButton("Generate Workbook")
        selection_layout.addWidget(QLabel("Main Location"), 0, 0)
        selection_layout.addWidget(self.cmb_building, 0, 1)
        selection_layout.addWidget(self.btn_preview_assets, 0, 2)
        selection_layout.addWidget(self.btn_preview_locations, 0, 3)
        selection_layout.addWidget(self.btn_mapping, 0, 4)
        selection_layout.addWidget(self.btn_generate, 0, 5)
        root.addWidget(selection)

        stats_layout = QHBoxLayout()
        self.asset_counter = QLabel("Assets: 0")
        self.room_counter = QLabel("Rooms: 0")
        self.duplicate_counter = QLabel("Duplicate assets: 0")
        self.missing_counter = QLabel("Missing sheets: 0")
        for label in (
            self.asset_counter,
            self.room_counter,
            self.duplicate_counter,
            self.missing_counter,
        ):
            label.setObjectName("counter")
            stats_layout.addWidget(label)
        root.addLayout(stats_layout)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        root.addWidget(self.progress)

        log_heading = QHBoxLayout()
        log_heading.addWidget(QLabel("Processing Log"))
        log_heading.addStretch(1)
        self.btn_export_log = QPushButton("Export Log")
        self.btn_open_missing = QPushButton("Open Missing Locations")
        self.btn_open_duplicates = QPushButton("Open Duplicate Report")
        log_heading.addWidget(self.btn_export_log)
        log_heading.addWidget(self.btn_open_missing)
        log_heading.addWidget(self.btn_open_duplicates)
        root.addLayout(log_heading)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        root.addWidget(self.log_view, 1)

        self.btn_browse_source.clicked.connect(self.browse_source)
        self.btn_browse_template.clicked.connect(self.browse_template)
        self.btn_browse_output.clicked.connect(self.browse_output)
        self.cmb_building.currentTextChanged.connect(self._update_counters)
        self.btn_preview_assets.clicked.connect(self.preview_assets)
        self.btn_preview_locations.clicked.connect(self.preview_locations)
        self.btn_mapping.clicked.connect(self.edit_location_mapping)
        self.btn_generate.clicked.connect(self.start_generation)
        self.btn_export_log.clicked.connect(self.export_log)
        self.btn_open_missing.clicked.connect(lambda: self._open_report("Missing_Locations.xlsx"))
        self.btn_open_duplicates.clicked.connect(lambda: self._open_report("Duplicate_Assets.xlsx"))

        self._restore_settings()

    def _restore_settings(self) -> None:
        self.source_path.setText(
            str(self.settings.value("source_path", "No Asset Monitoring workbook selected"))
        )
        self.template_path.setText(
            str(self.settings.value("template_path", str(DEFAULT_TEMPLATE)))
        )
        self.output_path.setText(
            str(self.settings.value("output_dir", str(DEFAULT_OUTPUT_DIR)))
        )

    def _persist_settings(self) -> None:
        self.settings.setValue("source_path", self.source_path.text())
        self.settings.setValue("template_path", self.template_path.text())
        self.settings.setValue("output_dir", self.output_path.text())

    def _log(self, message: str, level: str = "INFO") -> None:
        record = make_log_record(message, level)
        self.log_records.append(record)
        self.log_view.append(f"[{record['Timestamp']}] {level}: {message}")

    def browse_source(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Asset Monitoring Workbook",
            "",
            "Excel Workbooks (*.xlsx *.xlsm)",
        )
        if not path:
            return
        self.source_path.setText(path)
        self._persist_settings()
        self._log(f"Selected source workbook: {path}")
        self._start_worker(
            load_asset_register,
            path,
            on_success=self._source_loaded,
        )

    def browse_template(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Inventory Template",
            self.template_path.text(),
            "Excel Workbooks (*.xlsx)",
        )
        if path:
            self.template_path.setText(path)
            self._persist_settings()
            self._log(f"Selected template: {path}")

    def browse_output(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self,
            "Select Output Folder",
            self.output_path.text(),
        )
        if path:
            self.output_path.setText(path)
            self._persist_settings()

    def _start_worker(
        self,
        operation: Any,
        *args: Any,
        on_success: Any,
        **kwargs: Any,
    ) -> None:
        if self.worker is not None and self.worker.isRunning():
            QMessageBox.warning(self, "Busy", "Wait for the current operation to finish.")
            return
        self._set_busy(True)
        worker = OperationWorker(operation, *args, **kwargs)
        self.worker = worker
        worker.progress.connect(self._on_progress)
        worker.log_message.connect(self._log)
        worker.failed.connect(self._on_worker_failed)
        worker.succeeded.connect(on_success)
        worker.finished.connect(lambda: self._set_busy(False))
        worker.start()

    def _set_busy(self, busy: bool) -> None:
        for button in (
            self.btn_browse_source,
            self.btn_browse_template,
            self.btn_browse_output,
            self.btn_preview_assets,
            self.btn_preview_locations,
            self.btn_mapping,
            self.btn_generate,
            self.btn_export_log,
        ):
            button.setEnabled(not busy)
        self.cmb_building.setEnabled(not busy)
        if busy:
            self.progress.setValue(0)

    def _on_progress(self, value: int, message: str) -> None:
        self.progress.setValue(value)
        self.statusBar().showMessage(message)

    def _on_worker_failed(self, details: str) -> None:
        message = details.split("\n\n", 1)[0]
        self._log(message, "ERROR")
        QMessageBox.critical(self, "Operation Failed", message)

    def _source_loaded(self, result: LoadedAssets) -> None:
        self.loaded_assets = result
        self.location_overrides.clear()
        buildings = list_main_locations(
            result.dataframe,
            result.main_location_column,
        )
        self.cmb_building.clear()
        self.cmb_building.addItems(buildings)
        self._log(f"Loaded {len(result.dataframe):,} asset records from both registers")
        self._update_counters()

    def _selected_records(self) -> pd.DataFrame:
        if self.loaded_assets is None:
            return pd.DataFrame()
        return filter_by_main_location(
            self.loaded_assets.dataframe,
            self.loaded_assets.main_location_column,
            self.cmb_building.currentText(),
        )

    def _selected_rooms(self) -> list[str]:
        if self.loaded_assets is None:
            return []
        records = self._selected_records()
        mapper = LocationMapper(
            records[self.loaded_assets.specific_location_column].tolist(),
            self.location_overrides,
        )
        return mapper.locations

    def _update_counters(self) -> None:
        records = self._selected_records()
        self.asset_counter.setText(f"Assets: {len(records):,}")
        self.room_counter.setText(f"Rooms: {len(self._selected_rooms()):,}")

    def preview_assets(self) -> None:
        if self.loaded_assets is None:
            QMessageBox.information(self, "Preview", "Load an Asset Monitoring workbook first.")
            return
        records = self._selected_records()
        PreviewDialog(
            f"Asset Preview - {self.cmb_building.currentText()}",
            records,
            self,
            MAX_PREVIEW_ROWS,
        ).exec()

    def preview_locations(self) -> None:
        if self.loaded_assets is None:
            QMessageBox.information(self, "Preview", "Load an Asset Monitoring workbook first.")
            return
        PreviewDialog(
            f"Locations - {self.cmb_building.currentText()}",
            self._selected_rooms(),
            self,
        ).exec()

    def edit_location_mapping(self) -> None:
        if self.loaded_assets is None:
            QMessageBox.information(self, "Mapping", "Load an Asset Monitoring workbook first.")
            return
        values = self._selected_records()[
            self.loaded_assets.specific_location_column
        ].tolist()
        dialog = MappingDialog(values, self)
        if dialog.exec() == MappingDialog.DialogCode.Accepted:
            self.location_overrides = dialog.overrides()
            self._log(f"Applied {len(self.location_overrides)} location mapping(s)")
            self._update_counters()

    def start_generation(self) -> None:
        if self.loaded_assets is None:
            QMessageBox.warning(self, "Missing Source", "Browse and load an Asset Monitoring workbook first.")
            return
        selected_location = self.cmb_building.currentText().strip()
        if not selected_location:
            QMessageBox.warning(self, "Missing Location", "Select a Main Location.")
            return
        self._persist_settings()
        self.log_records = []
        self._log(f"Starting workbook generation for {selected_location}")

        loaded_assets = self.loaded_assets
        template_path = self.template_path.text()
        output_dir = self.output_path.text()
        location_overrides = dict(self.location_overrides)
        self._start_worker(
            generate_workbook,
            loaded_assets.dataframe,
            selected_location,
            template_path,
            output_dir,
            main_location_column=loaded_assets.main_location_column,
            specific_location_column=loaded_assets.specific_location_column,
            location_overrides=location_overrides,
            on_success=self._generation_completed,
        )

    def _generation_completed(self, result: GenerationResult) -> None:
        self.last_result = result
        self.asset_counter.setText(f"Assets: {result.asset_count:,}")
        self.room_counter.setText(f"Rooms: {result.room_count:,}")
        self.duplicate_counter.setText(f"Duplicate assets: {result.duplicate_count:,}")
        self.missing_counter.setText(f"Missing sheets: {len(result.missing_locations):,}")
        self._log(f"Created: {result.workbook_path}")
        self._log(f"Reports saved under: {self.output_path.text()}")
        QMessageBox.information(
            self,
            "Generation Complete",
            f"Workbook created:\n{result.workbook_path}\n\n"
            f"Assets: {result.asset_count:,}\n"
            f"Location sheets: {result.room_count:,}\n"
            f"Duplicate assets: {result.duplicate_count:,}",
        )

    def export_log(self) -> None:
        try:
            path = export_processing_log(self.output_path.text(), self.log_records)
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", f"Could not export the processing log:\n{exc}")
            return
        self._log(f"Processing log exported: {path}")
        QMessageBox.information(self, "Log Exported", f"Saved to:\n{path}")

    def _open_report(self, filename: str) -> None:
        path = Path(self.output_path.text()) / filename
        if not path.is_file():
            QMessageBox.information(self, "Report Not Found", f"Generate a workbook first:\n{path}")
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(path))):
            QMessageBox.warning(self, "Could Not Open", f"Could not open report:\n{path}")

    def closeEvent(self, event: Any) -> None:
        if self.worker is not None and self.worker.isRunning():
            QMessageBox.warning(
                self,
                "Operation in Progress",
                "Wait for the current workbook operation to finish before closing.",
            )
            event.ignore()
            return
        self._persist_settings()
        event.accept()
