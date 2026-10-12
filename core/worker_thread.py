from __future__ import annotations

import traceback
from collections.abc import Callable
from inspect import signature
from typing import Any

from PySide6.QtCore import QThread, Signal


class OperationWorker(QThread):
    succeeded = Signal(object)
    failed = Signal(str)
    progress = Signal(int, str)
    log_message = Signal(str)

    def __init__(
        self,
        operation: Callable[..., Any],
        *args: Any,
        **kwargs: Any,
    ) -> None:
        super().__init__()
        self._operation = operation
        self._args = args
        self._kwargs = kwargs

    def run(self) -> None:
        try:
            kwargs = dict(self._kwargs)
            parameters = signature(self._operation).parameters
            if "progress_callback" in parameters:
                kwargs.setdefault("progress_callback", self.progress.emit)
            if "log_callback" in parameters:
                kwargs.setdefault("log_callback", self.log_message.emit)
            result = self._operation(*self._args, **kwargs)
        except Exception as exc:
            self.failed.emit(f"{exc}\n\n{traceback.format_exc()}")
        else:
            self.succeeded.emit(result)
