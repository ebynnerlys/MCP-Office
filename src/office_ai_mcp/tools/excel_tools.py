from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from office_ai_mcp.models.requests import ExcelRangeRequest, ExcelWriteRangeRequest, ExportPdfRequest
from office_ai_mcp.services.excel_service import ExcelService


def register_excel_tools(mcp: FastMCP, service: ExcelService) -> None:
    @mcp.tool(
        name="excel_list_sheets",
        description="List worksheets in an Excel workbook with their used row and column counts.",
    )
    def excel_list_sheets(path: str) -> dict[str, object]:
        """List worksheet names and basic used-range dimensions in a workbook."""
        return service.list_sheets(path).model_dump()

    @mcp.tool(
        name="excel_read_range",
        description="Read values from a specific worksheet range in an Excel workbook.",
    )
    def excel_read_range(path: str, sheet: str, cell_range: str) -> dict[str, object]:
        """Read the values contained in a worksheet range."""
        request = ExcelRangeRequest(path=path, sheet=sheet, cell_range=cell_range, create_backup=False)
        return service.read_range(
            path=request.path,
            sheet=request.sheet,
            cell_range=request.cell_range,
        ).model_dump()

    @mcp.tool(
        name="excel_write_range",
        description="Write values into a worksheet range in an Excel workbook, optionally creating a backup first.",
    )
    def excel_write_range(
        path: str,
        sheet: str,
        cell_range: str,
        values: Any,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Write one value or a matrix of values into a worksheet range."""
        request = ExcelWriteRangeRequest(
            path=path,
            sheet=sheet,
            cell_range=cell_range,
            values=values,
            create_backup=create_backup,
        )
        return service.write_range(
            path=request.path,
            sheet=request.sheet,
            cell_range=request.cell_range,
            values=request.values,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="excel_export_pdf",
        description="Export an Excel workbook to PDF without changing workbook contents.",
    )
    def excel_export_pdf(path: str, out_path: str) -> dict[str, object]:
        """Export a workbook to a PDF file."""
        request = ExportPdfRequest(path=path, out_path=out_path)
        return service.export_pdf(path=request.path, out_path=request.out_path).model_dump()
