"""Regel-Registry – sammelt alle verfügbaren Regeln."""

from excel_checker.rules.base import BaseRule
from excel_checker.rules.structure import (
    MergedCellsRule,
    DataTypeHomogeneityRule,
    HeaderDetectionRule,
    EmptyRowColSeparatorRule,
    IdentifierConsistencyRule,
    MissingUniqueKeyRule,
    TextBasedIdRule,
)
from excel_checker.rules.formulas import (
    AbsoluteVsRelativeRule,
    VolatileFunctionsRule,
    VlookupChainsRule,
    CircularReferenceHintRule,
    CrossSheetReferenceRule,
)
from excel_checker.rules.formula_errors import FormulaErrorValuesRule
from excel_checker.rules.volume import (
    RowCountRule,
    FormulaDensityRule,
    SheetCountRule,
    FileSizeRule,
)
from excel_checker.rules.implicit_knowledge import (
    UndocumentedColorCodesRule,
    HiddenSheetsRule,
    HiddenRowsColumnsRule,
    ConditionalFormattingRule,
    MagicNumbersRule,
    DataValidationRule,
    CrypticSheetNamesRule,
    CommentsWithLogicRule,
    CustomNumberFormatsRule,
    ProtectedAreasRule,
)

__all__ = ["ALL_RULES", "BaseRule"]

ALL_RULES: list[type[BaseRule]] = [
    # Struktur & Normalform
    MergedCellsRule,
    DataTypeHomogeneityRule,
    HeaderDetectionRule,
    EmptyRowColSeparatorRule,
    IdentifierConsistencyRule,
    MissingUniqueKeyRule,
    TextBasedIdRule,
    # Formeln & Bezüge
    AbsoluteVsRelativeRule,
    VolatileFunctionsRule,
    VlookupChainsRule,
    CircularReferenceHintRule,
    CrossSheetReferenceRule,
    FormulaErrorValuesRule,
    # Volumen & Limits
    RowCountRule,
    FormulaDensityRule,
    SheetCountRule,
    FileSizeRule,
    # Implizites Wissen
    UndocumentedColorCodesRule,
    HiddenSheetsRule,
    HiddenRowsColumnsRule,
    ConditionalFormattingRule,
    MagicNumbersRule,
    DataValidationRule,
    CrypticSheetNamesRule,
    CommentsWithLogicRule,
    CustomNumberFormatsRule,
    ProtectedAreasRule,
]
