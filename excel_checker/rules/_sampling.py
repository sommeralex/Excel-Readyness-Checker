"""Sampling-Framework für Regeln, die bei großen Dateien nur eine Stichprobe prüfen.

Wird vom Engine aktiviert, wenn die Datei die Schwelle für Tier 2/3 überschreitet.
Regeln mit ``supports_sampling = True`` respektieren ``SampleMode`` und ergänzen
Findings mit einem ``sample_note``, damit im Report erkennbar ist, dass nur eine
Stichprobe geprüft wurde.
"""

from __future__ import annotations

import random
from dataclasses import dataclass
from typing import Iterator, Optional


@dataclass(frozen=True)
class SampleMode:
    """Steuert Stichproben-Verhalten für cell-iterierende Regeln."""

    max_cells_per_sheet: int = 10_000
    max_rows_per_sheet: int = 20_000
    max_cols_per_sheet: int = 300
    seed: int = 42

    def disclosure_de(self) -> str:
        return (
            f"Stichprobe: max. {self.max_cells_per_sheet:,} Zellen pro Blatt"
        ).replace(",", ".")

    def disclosure_en(self) -> str:
        return f"Sample: up to {self.max_cells_per_sheet:,} cells per sheet"


def bounded_range(total: int, cap: int, seed: int) -> list[int]:
    """Gibt eine Liste von 1-basierten Indizes zurück, die eine Stichprobe
    über den Bereich ``1..total`` bilden. Bei ``total <= cap`` wird die
    vollständige Reihenfolge ``[1..total]`` zurückgegeben. Sonst wird
    eine reproduzierbare Stichprobe der Größe ``cap`` gezogen und sortiert.
    """
    if total <= 0:
        return []
    if total <= cap:
        return list(range(1, total + 1))
    rng = random.Random(seed)
    sampled = rng.sample(range(1, total + 1), cap)
    sampled.sort()
    return sampled


def iter_sampled_rows(
    ws,
    max_row: int,
    max_col: int,
    sample_mode: Optional[SampleMode],
) -> Iterator[tuple]:
    """Liefert Zeilen des Sheets, ggf. als Stichprobe.

    Wenn ``sample_mode is None`` → iteriert alle Zeilen bis ``max_row``.
    Sonst: cappt Zeilen und Spalten auf die Grenzen aus ``SampleMode``
    und zieht bei Bedarf eine reproduzierbare Stichprobe.

    Yields Tupel ``(row_idx, row_cells)`` wobei ``row_cells`` ein Tuple
    von Cell-Objekten ist (analog zu openpyxl ``iter_rows``).
    """
    if max_row <= 0 or max_col <= 0:
        return

    if sample_mode is None:
        for row_idx, row in enumerate(
            ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col), start=1
        ):
            yield row_idx, row
        return

    effective_cols = min(max_col, sample_mode.max_cols_per_sheet)
    row_indices = bounded_range(max_row, sample_mode.max_rows_per_sheet, sample_mode.seed)

    cell_budget = sample_mode.max_cells_per_sheet
    cells_yielded = 0

    for row_idx in row_indices:
        if cells_yielded >= cell_budget:
            break
        row_tuple = next(
            ws.iter_rows(min_row=row_idx, max_row=row_idx, max_col=effective_cols),
            None,
        )
        if row_tuple is None:
            continue
        yield row_idx, row_tuple
        cells_yielded += effective_cols
