"""addConditionalFormat — apply conditional formatting rules (cellValue, formula, colorScale, dataBar, iconSet, text) to a range."""

from __future__ import annotations

from typing import Any, Optional

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


# Excel XlFormatConditionType enum values.
_CF_TYPE = {
    "cellValue": 1,    # xlCellValue
    "formula": 2,      # xlExpression
    "colorScale": 3,   # xlColorScale
    "dataBar": 4,      # xlDatabar
    "iconSet": 6,      # xlIconSets
    "text": 9,         # xlTextString
}

# Excel XlFormatConditionOperator for cellValue rules.
_CV_OPERATOR = {
    "between": 1,               # xlBetween
    "notBetween": 2,            # xlNotBetween
    "equalTo": 3,               # xlEqual
    "notEqualTo": 4,            # xlNotEqual
    "greaterThan": 5,           # xlGreater
    "lessThan": 6,              # xlLess
    "greaterThanOrEqualTo": 7,  # xlGreaterEqual
    "lessThanOrEqualTo": 8,     # xlLessEqual
}

# xlTextString operator default — xlContains = 0 (TextOperator.xlContains)
_TEXT_CONTAINS = 0


def _hex_to_rgb_int(hex_color: str) -> int:
    """Convert "#RRGGBB" → Excel long int (BGR packed)."""
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    return r + g * 256 + b * 65536


def _apply_format(fc: Any, fmt: dict[str, Any]) -> None:
    """Apply fill/font/bold formatting from the params' `format` block."""
    if not fmt:
        return
    fill_color = fmt.get("fillColor")
    if fill_color:
        try:
            fc.Interior.Color = _hex_to_rgb_int(fill_color)
        except Exception:
            pass
    font_color = fmt.get("fontColor")
    if font_color:
        try:
            fc.Font.Color = _hex_to_rgb_int(font_color)
        except Exception:
            pass
    bold = fmt.get("bold")
    if bold is not None:
        try:
            fc.Font.Bold = bool(bold)
        except Exception:
            pass


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    rule_type = params.get("ruleType")
    if not address or not rule_type:
        return {"status": "error", "message": "addConditionalFormat requires 'range' and 'ruleType'."}

    if rule_type not in _CF_TYPE:
        return {
            "status": "error",
            "message": f"Unsupported ruleType {rule_type!r}. Expected one of {sorted(_CF_TYPE)}.",
        }

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would add {rule_type} conditional format to {rng.address}.",
        }

    try:
        fmt_conditions = rng.api.FormatConditions
        fmt = params.get("format") or {}
        values = params.get("values") or []
        operator_key = params.get("operator") or "greaterThan"
        text = params.get("text")
        formula = params.get("formula")

        if rule_type == "cellValue":
            op = _CV_OPERATOR.get(operator_key, _CV_OPERATOR["greaterThan"])
            f1 = str(values[0]) if len(values) >= 1 else "0"
            f2: Optional[str] = str(values[1]) if len(values) >= 2 else None
            if f2 is not None:
                cf = fmt_conditions.Add(Type=_CF_TYPE["cellValue"], Operator=op, Formula1=f1, Formula2=f2)
            else:
                cf = fmt_conditions.Add(Type=_CF_TYPE["cellValue"], Operator=op, Formula1=f1)
            _apply_format(cf, fmt)

        elif rule_type == "formula":
            formula_str = formula or text or "=TRUE"
            cf = fmt_conditions.Add(Type=_CF_TYPE["formula"], Formula1=formula_str)
            _apply_format(cf, fmt)

        elif rule_type == "colorScale":
            # Add a 3-color scale: red (low) → yellow (mid) → green (high).
            cf = fmt_conditions.Add(Type=_CF_TYPE["colorScale"])
            try:
                # ColorScaleCriteria is a 1-based collection of 3 stops.
                # Type enum: xlConditionValueLowestValue=1, xlConditionValuePercent=3,
                #            xlConditionValueHighestValue=2
                cf.ColorScaleCriteria(1).Type = 1  # lowest
                cf.ColorScaleCriteria(1).FormatColor.Color = _hex_to_rgb_int("#f4cccc")
                cf.ColorScaleCriteria(2).Type = 3  # percent
                cf.ColorScaleCriteria(2).Value = 50
                cf.ColorScaleCriteria(2).FormatColor.Color = _hex_to_rgb_int("#fff2cc")
                cf.ColorScaleCriteria(3).Type = 2  # highest
                cf.ColorScaleCriteria(3).FormatColor.Color = _hex_to_rgb_int("#d9ead3")
            except Exception:
                # Some COM builds expose the criteria differently; leave defaults if so.
                pass

        elif rule_type == "dataBar":
            fmt_conditions.Add(Type=_CF_TYPE["dataBar"])

        elif rule_type == "iconSet":
            # xlIconSets=6. The default icon set is 3-traffic-lights-unrimmed.
            fmt_conditions.Add(Type=_CF_TYPE["iconSet"])

        elif rule_type == "text":
            if not text:
                return {"status": "error", "message": "addConditionalFormat text rule requires 'text'."}
            # xlTextString uses Operator=xlContains (0) and String=text.
            try:
                cf = fmt_conditions.Add(
                    Type=_CF_TYPE["text"], Operator=_TEXT_CONTAINS, String=text
                )
            except Exception:
                # Fallback: some COM bindings require the text via Formula1.
                cf = fmt_conditions.Add(
                    Type=_CF_TYPE["text"], Operator=_TEXT_CONTAINS, Formula1=text
                )
            _apply_format(cf, fmt)

    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Conditional format failed: {exc}",
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"Applied {rule_type} conditional format to {rng.address}.",
        "outputs": {"range": rng.address, "ruleType": rule_type},
    }


registry.register("addConditionalFormat", handler, mutates=True, affects_formatting=True)
