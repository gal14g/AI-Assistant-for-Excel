"""pageLayout — configure sheet page layout (margins, orientation, paper size, print area, gridlines)."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


# XlPageOrientation: xlPortrait=1, xlLandscape=2.
_ORIENTATION = {"portrait": 1, "landscape": 2}

# XlPaperSize common values.
_PAPER = {
    "letter": 1,     # xlPaperLetter
    "legal": 5,      # xlPaperLegal
    "a3": 8,         # xlPaperA3
    "a4": 9,         # xlPaperA4
    "a5": 11,        # xlPaperA5
    "b4": 12,        # xlPaperB4
    "b5": 13,        # xlPaperB5
    "tabloid": 3,    # xlPaperTabloid (11x17)
}


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    sheet_name = params.get("sheetName")
    margins = params.get("margins")
    orientation = params.get("orientation")
    paper_size = params.get("paperSize")
    print_area = params.get("printArea")
    show_gridlines = params.get("showGridlines")
    print_gridlines = params.get("printGridlines")

    book = ctx.workbook_handle

    if ctx.dry_run:
        parts: list[str] = []
        if margins:
            parts.append("margins")
        if orientation:
            parts.append(f"orientation={orientation}")
        if paper_size:
            parts.append(f"paperSize={paper_size}")
        if print_area:
            parts.append(f"printArea={print_area}")
        if show_gridlines is not None:
            parts.append(f"showGridlines={show_gridlines}")
        if print_gridlines is not None:
            parts.append(f"printGridlines={print_gridlines}")
        return {
            "status": "preview",
            "message": f"Would set page layout: {', '.join(parts) or 'no-op'}.",
        }

    # Resolve the target sheet.
    if sheet_name:
        try:
            sheet = book.sheets[sheet_name]
        except Exception:
            return {
                "status": "error",
                "message": f"Sheet {sheet_name!r} not found. Please check the sheet name.",
            }
    else:
        sheet = book.sheets.active

    try:
        page_setup = sheet.api.PageSetup

        # Margins — TS multiplies inches by 72 (points-per-inch). Match that.
        if margins:
            if margins.get("top") is not None:
                page_setup.TopMargin = margins["top"] * 72
            if margins.get("bottom") is not None:
                page_setup.BottomMargin = margins["bottom"] * 72
            if margins.get("left") is not None:
                page_setup.LeftMargin = margins["left"] * 72
            if margins.get("right") is not None:
                page_setup.RightMargin = margins["right"] * 72
            if margins.get("header") is not None:
                page_setup.HeaderMargin = margins["header"] * 72
            if margins.get("footer") is not None:
                page_setup.FooterMargin = margins["footer"] * 72

        # Orientation.
        if orientation:
            value = _ORIENTATION.get(orientation.lower(), _ORIENTATION["portrait"])
            page_setup.Orientation = value

        # Paper size.
        if paper_size:
            paper = _PAPER.get(paper_size.lower())
            if paper is not None:
                page_setup.PaperSize = paper

        # Print area.
        if print_area:
            page_setup.PrintArea = print_area

        # Gridlines — showGridlines affects the display (ActiveWindow), printGridlines
        # is a print-time page setup flag.
        if show_gridlines is not None:
            try:
                sheet.activate()
                book.app.api.ActiveWindow.DisplayGridlines = bool(show_gridlines)
            except Exception:
                pass
        if print_gridlines is not None:
            page_setup.PrintGridlines = bool(print_gridlines)

    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Page layout failed: {exc}",
            "error": str(exc),
        }

    changes: list[str] = []
    if margins:
        changes.append("margins")
    if orientation:
        changes.append(orientation)
    if paper_size:
        changes.append(paper_size)
    if print_area:
        changes.append(f"print area {print_area}")
    if show_gridlines is not None:
        changes.append(f"gridlines {'on' if show_gridlines else 'off'}")
    if print_gridlines is not None:
        changes.append(f"print gridlines {'on' if print_gridlines else 'off'}")

    suffix = f' on "{sheet_name}"' if sheet_name else ""
    return {
        "status": "success",
        "message": f"Page layout updated: {', '.join(changes) or 'no-op'}{suffix}.",
        "outputs": {"sheet": sheet.name},
    }


registry.register("pageLayout", handler, mutates=False, affects_formatting=True)
