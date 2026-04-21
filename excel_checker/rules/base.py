"""Basis-Klasse für alle Prüfregeln (Rule Engine)."""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import List

import openpyxl

from excel_checker.models import Finding


class BaseRule(ABC):
    """Jede Regel erbt von dieser Klasse."""

    @property
    @abstractmethod
    def rule_id(self) -> str:
        """Eindeutige Regel-ID, z. B. 'STR-001'."""

    @property
    @abstractmethod
    def rule_name(self) -> str:
        """Menschenlesbarer Name der Regel."""

    def check(self, workbook: openpyxl.Workbook, file_path: str, progress_callback: callable = None) -> List[Finding]:
        """Führt die Prüfung aus und gibt Findings zurück. Kann optional einen Fortschritts-Callback nutzen."""
        raise NotImplementedError()
