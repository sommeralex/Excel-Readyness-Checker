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


def iter_window_rows(ws, max_row: int, max_col: int, min_row: int = 1) -> Iterator[tuple]:
    """Vorwaerts-Iteration ueber ein begrenztes Fenster des Blatts.

    Ersatz fuer ``ws.cell(row=r, column=c)``: im ``read_only``-Modus kann
    openpyxl das Blatt nur streamen, wahlfreier Zellzugriff ist dort nicht
    moeglich (und ohne ``read_only`` quadratisch teuer). Da der Generator
    nach ``max_row`` abbricht, wird bei grossen Dateien nur das Fenster
    geparst, nicht das ganze Blatt.

    Yields ``(row_idx, row_cells)``. ``row_idx`` ist die echte 1-basierte
    Zeilennummer, ``row_cells`` ist ein auf ``max_col`` aufgefuelltes Tupel —
    Position ``i`` entspricht also Spalte ``i + 1``, auch bei Luecken.

    Achtung: Fuellzellen sind ``EmptyCell`` und haben im read-only-Modus
    weder ``.row`` noch ``.column``. Spaltennummern immer aus dem Index
    ableiten, nie aus dem Zellobjekt.
    """
    if max_row < min_row or max_col <= 0:
        return
    for offset, row in enumerate(
        ws.iter_rows(min_row=min_row, max_row=max_row, max_col=max_col)
    ):
        yield min_row + offset, row


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
        yield from iter_window_rows(ws, max_row, max_col)
        return

    effective_cols = min(max_col, sample_mode.max_cols_per_sheet)
    wanted = set(
        bounded_range(max_row, sample_mode.max_rows_per_sheet, sample_mode.seed)
    )
    if not wanted:
        return

    cell_budget = sample_mode.max_cells_per_sheet
    cells_yielded = 0

    # Ein einziger Vorwaertslauf bis zur letzten gezogenen Zeile. Frueher wurde
    # pro Stichprobenzeile ein eigener ``iter_rows``-Aufruf gemacht; unter
    # ``read_only`` haette das die Blatt-XML je Zeile neu geparst.
    for row_idx, row_tuple in iter_window_rows(ws, max(wanted), effective_cols):
        if row_idx not in wanted:
            continue
        if cells_yielded >= cell_budget:
            break
        yield row_idx, row_tuple
        cells_yielded += effective_cols
