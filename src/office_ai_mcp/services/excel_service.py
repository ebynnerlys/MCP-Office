from __future__ import annotations

from contextlib import suppress
from typing import Any

from tenacity import retry, stop_after_attempt, wait_fixed

from office_ai_mcp.config import Settings
from office_ai_mcp.models.responses import ExcelRangeResult, OperationResult, WorkbookSheetsResult, WorksheetSummary
from office_ai_mcp.services.base import OfficeService
from office_ai_mcp.utils.com_cleanup import office_application


def _normalize_excel_value(value: Any) -> Any:
    if isinstance(value, tuple):
        return [_normalize_excel_value(item) for item in value]
    return value


def _to_excel_value(value: Any) -> Any:
    if isinstance(value, list):
        return tuple(_to_excel_value(item) for item in value)
    return value


class ExcelService(OfficeService):
    allowed_suffixes = (".xls", ".xlsx", ".xlsm")

    def __init__(self, settings: Settings) -> None:
        super().__init__(settings)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def list_sheets(self, path: str) -> WorkbookSheetsResult:
        source = self.resolve_document_path(path)

        with office_application("Excel.Application", visible=self.settings.office_visible) as excel:
            workbook = None
            try:
                workbook = excel.Workbooks.Open(str(source), ReadOnly=True)
                sheets: list[WorksheetSummary] = []

                for index in range(1, int(workbook.Worksheets.Count) + 1):
                    worksheet = workbook.Worksheets(index)
                    used_range = worksheet.UsedRange
                    sheets.append(
                        WorksheetSummary(
                            name=str(worksheet.Name),
                            rows=int(used_range.Rows.Count),
                            columns=int(used_range.Columns.Count),
                        )
                    )

                return WorkbookSheetsResult(file_path=str(source), sheets=sheets)
            finally:
                if workbook is not None:
                    with suppress(Exception):
                        workbook.Close(SaveChanges=False)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def read_range(self, path: str, sheet: str, cell_range: str) -> ExcelRangeResult:
        source = self.resolve_document_path(path)

        with office_application("Excel.Application", visible=self.settings.office_visible) as excel:
            workbook = None
            try:
                workbook = excel.Workbooks.Open(str(source), ReadOnly=True)
                worksheet = workbook.Worksheets(sheet)
                values = worksheet.Range(cell_range).Value
                return ExcelRangeResult(
                    file_path=str(source),
                    sheet=sheet,
                    cell_range=cell_range,
                    values=_normalize_excel_value(values),
                )
            finally:
                if workbook is not None:
                    with suppress(Exception):
                        workbook.Close(SaveChanges=False)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def write_range(
        self,
        path: str,
        sheet: str,
        cell_range: str,
        values: Any,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with office_application("Excel.Application", visible=self.settings.office_visible) as excel:
            workbook = None
            try:
                workbook = excel.Workbooks.Open(str(source), ReadOnly=False)
                worksheet = workbook.Worksheets(sheet)
                worksheet.Range(cell_range).Value = _to_excel_value(values)
                workbook.Save()
                return OperationResult(
                    message="Excel range updated",
                    file_path=str(source),
                    backup_path=backup_path,
                    details={"sheet": sheet, "cell_range": cell_range},
                )
            finally:
                if workbook is not None:
                    with suppress(Exception):
                        workbook.Close(SaveChanges=False)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def export_pdf(self, path: str, out_path: str) -> OperationResult:
        source = self.resolve_document_path(path)
        target = self.resolve_output_path(out_path, allowed_suffixes=(".pdf",))

        with office_application("Excel.Application", visible=self.settings.office_visible) as excel:
            workbook = None
            try:
                workbook = excel.Workbooks.Open(str(source), ReadOnly=True)
                workbook.ExportAsFixedFormat(0, str(target))
                return OperationResult(
                    message="Excel workbook exported to PDF",
                    file_path=str(source),
                    details={"out_path": str(target)},
                )
            finally:
                if workbook is not None:
                    with suppress(Exception):
                        workbook.Close(SaveChanges=False)
