"""Basis-Klasse für alle Prüfregeln (Rule Engine)."""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Callable, List, Optional

import openpyxl

from excel_checker.models import Finding
from excel_checker.rules._sampling import SampleMode


class BaseRule(ABC):
    """Jede Regel erbt von dieser Klasse.

    Zusätzlich zu ``rule_id`` und ``rule_name`` klassifizieren drei
    Klassen-Attribute, wie sich die Regel unter Last verhält. Der Engine
    nutzt diese Attribute, um bei großen Dateien passende Regeln
    auszuwählen und ggf. Sampling zu aktivieren.

    - ``needs_styles``: Greift auf ``cell.fill``/``cell.font``/``cell.comment``
      oder auf Blatt-Eigenschaften wie ``ws.merged_cells``, ``ws.protection``,
      ``ws.conditional_formatting``, ``ws.data_validations`` zu. Solche Regeln
      können NICHT im read-only-Modus von openpyxl laufen — die Attribute
      existieren dort nicht und der Zugriff wirft ``AttributeError``.
    - ``scales_with_cells``: Laufzeit wächst linear mit Zellenzahl (iteriert
      alle Zellen). Kandidat für Sampling.
    - ``supports_sampling``: Regel darf mit einer ``SampleMode``-Instanz
      aufgerufen werden und liefert auch auf einer Stichprobe sinnvolle
      Ergebnisse (z. B. statistische Muster). Regeln, die jedes Vorkommen
      finden MÜSSEN (Fehlerwerte, Zirkelbezüge), setzen das auf False.
    """

    # Default-Werte sind konservativ: lieber eine Regel weglassen als
    # falsche/unvollständige Findings auf großen Dateien liefern.
    needs_styles: bool = False
    scales_with_cells: bool = False
    supports_sampling: bool = False

    # Vom Engine vor jedem ``check`` gesetzt. Nötig, weil ``file_path`` bei
    # einer Analyse aus dem Speicher (Browser) nur ein Anzeigename ist und
    # sich nicht vermessen lässt.
    file_size_bytes: Optional[int] = None

    @property
    @abstractmethod
    def rule_id(self) -> str:
        """Eindeutige Regel-ID, z. B. 'STR-001'."""

    @property
    @abstractmethod
    def rule_name(self) -> str:
        """Menschenlesbarer Name der Regel."""

    def check(
        self,
        workbook: openpyxl.Workbook,
        file_path: str,
        progress_callback: Optional[Callable] = None,
        sample_mode: Optional[SampleMode] = None,
    ) -> List[Finding]:
        """Führt die Prüfung aus und gibt Findings zurück.

        ``sample_mode`` wird vom Engine nur an Regeln übergeben, die
        ``supports_sampling = True`` gesetzt haben. Regeln müssen die
        resultierenden Findings mit ``sample_note`` markieren, damit der
        Report transparent macht, dass nur eine Stichprobe geprüft wurde.
        """
        raise NotImplementedError()
