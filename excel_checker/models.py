"""Datenmodelle für Regelergebnisse und Reports."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional


class Severity(Enum):
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    CRITICAL = "critical"


class Category(Enum):
    STRUCTURE = "structure"
    FORMULA = "formula"
    VOLUME = "volume"

    @property
    def label(self) -> str:
        from excel_checker.i18n import _
        return _(f"cat.{self.value}")


_RECTYPE_I18N_KEYS = {
    "DATABASE_MIGRATION": "rectype.db_migration",
    "POWER_BI": "rectype.power_bi",
    "DATA_WAREHOUSE": "rectype.data_warehouse",
    "SHAREPOINT_LIST": "rectype.sharepoint_list",
    "NORMALIZATION": "rectype.normalization",
    "SPLIT_WORKBOOK": "rectype.split_workbook",
    "CLEANUP": "rectype.cleanup",
    "OK": "rectype.ok",
}


class RecommendationType(Enum):
    DATABASE_MIGRATION = "db_migration"
    POWER_BI = "power_bi"
    DATA_WAREHOUSE = "data_warehouse"
    SHAREPOINT_LIST = "sharepoint_list"
    NORMALIZATION = "normalization"
    SPLIT_WORKBOOK = "split_workbook"
    CLEANUP = "cleanup"
    OK = "ok"

    @property
    def label(self) -> str:
        from excel_checker.i18n import _
        return _(_RECTYPE_I18N_KEYS[self.name])


@dataclass
class Recommendation:
    """Eine strategische Empfehlung für den professionellen Umgang."""

    rec_type: RecommendationType
    priority: int  # 1 = höchste Priorität
    title: str
    reason: str
    action: str  # Konkrete Handlungsempfehlung


@dataclass
class Finding:
    """Ein einzelnes Prüfergebnis."""

    rule_id: str
    rule_name: str
    category: Category
    severity: Severity
    message: str
    sheet: Optional[str] = None
    cell: Optional[str] = None
    detail: Optional[str] = None
    suggestion: Optional[str] = None  # Konkreter Verbesserungsvorschlag
    score_penalty: int = 0  # Abzug vom Health-Score (0-100)


@dataclass
class SheetStats:
    """Statistische Kennzahlen eines Tabellenblatts."""

    name: str
    row_count: int = 0
    col_count: int = 0
    formula_count: int = 0
    static_count: int = 0
    merged_regions: int = 0
    empty_row_gaps: int = 0
    empty_col_gaps: int = 0
    db_readiness: int = 0  # DB-Migrationstauglichkeit (0-100)
    is_phantom: bool = False  # Phantom-Dimensionen erkannt (max_row = 1M)


@dataclass
class WorkbookReport:
    """Gesamtbericht über eine Excel-Datei."""

    file_path: str
    file_size_mb: float = 0.0
    sheet_count: int = 0
    findings: List[Finding] = field(default_factory=list)
    sheet_stats: List[SheetStats] = field(default_factory=list)
    recommendations: List[Recommendation] = field(default_factory=list)

    @property
    def health_score(self) -> int:
        """Berechnet den Health-Score (0-100). 100 = perfekt."""
        total_penalty = sum(f.score_penalty for f in self.findings)
        return max(0, 100 - total_penalty)

    @property
    def critical_count(self) -> int:
        return sum(1 for f in self.findings if f.severity == Severity.CRITICAL)

    @property
    def error_count(self) -> int:
        return sum(1 for f in self.findings if f.severity == Severity.ERROR)

    @property
    def warning_count(self) -> int:
        return sum(1 for f in self.findings if f.severity == Severity.WARNING)

    @property
    def info_count(self) -> int:
        return sum(1 for f in self.findings if f.severity == Severity.INFO)
