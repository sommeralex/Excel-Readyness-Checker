"""Regel: Fehlerwerte in Formeln (#DIV/0!, #WERT!, #BEZUG!, #NAME?)"""

from __future__ import annotations
from typing import List
import openpyxl
from excel_checker.models import Category, Finding, Severity
from excel_checker.rules.base import BaseRule
from excel_checker.i18n import _

ERROR_VALUES = {"#DIV/0!", "#WERT!", "#BEZUG!", "#NAME?", "#NULL!", "#ZAHL!", "#NV", "#REF!", "#VALUE!", "#NUM!", "#N/A"}

class FormulaErrorValuesRule(BaseRule):
    """Prüft auf Fehlerwerte in Formeln (z.B. #DIV/0!, #WERT!, #BEZUG!, #NAME?)."""
    rule_id = "FRM-006"
    rule_name = _("FRM-006.name")

    def check(self, workbook: openpyxl.Workbook, file_path: str) -> List[Finding]:
        findings = []
        for ws in workbook.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type == 'e' and str(cell.value).upper() in ERROR_VALUES:
                        findings.append(Finding(
                            rule_id=self.rule_id,
                            rule_name=self.rule_name,
                            category=Category.FORMULA,
                            severity=Severity.ERROR,
                            message=_("FRM-006.msg", error=str(cell.value)),
                            sheet=ws.title,
                            cell=cell.coordinate,
                            detail=_("FRM-006.detail", error=str(cell.value)),
                            suggestion=_("FRM-006.tip"),
                            score_penalty=7,
                        ))
        return findings
