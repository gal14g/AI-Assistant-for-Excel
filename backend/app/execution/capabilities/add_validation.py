"""addValidation — apply data validation rules (list, wholeNumber, decimal, date, textLength, custom) to a range."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


# Excel XlDVType enum values.
_V_TYPE = {
    "list": 3,          # xlValidateList
    "wholeNumber": 1,   # xlValidateWholeNumber
    "decimal": 2,       # xlValidateDecimal
    "date": 4,          # xlValidateDate
    "textLength": 6,    # xlValidateTextLength
    "custom": 7,        # xlValidateCustom
}

# Excel XlFormatConditionOperator mapping for validation.
_V_OPERATOR = {
    "between": 1,               # xlBetween
    "notBetween": 2,            # xlNotBetween
    "equalTo": 3,               # xlEqual
    "notEqualTo": 4,            # xlNotEqual
    "greaterThan": 5,           # xlGreater
    "lessThan": 6,              # xlLess
    "greaterThanOrEqualTo": 7,  # xlGreaterEqual
    "lessThanOrEqualTo": 8,     # xlLessEqual
}

# XlDVAlertStyle — xlValidAlertStop = 1 (matches the TS default).
_ALERT_STOP = 1


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    validation_type = params.get("validationType")
    if not address or not validation_type:
        return {"status": "error", "message": "addValidation requires 'range' and 'validationType'."}

    if validation_type not in _V_TYPE:
        return {
            "status": "error",
            "message": f"Unsupported validationType {validation_type!r}. Expected one of {sorted(_V_TYPE)}.",
        }

    list_values = params.get("listValues")
    operator_key = params.get("operator") or "between"
    minv = params.get("min")
    maxv = params.get("max")
    formula = params.get("formula")
    show_error_alert = params.get("showErrorAlert", True)
    error_message = params.get("errorMessage") or "The value you entered is not valid."

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would add {validation_type} validation to {rng.address}.",
        }

    try:
        v = rng.api.Validation
        # Clear any existing rule first — Validation.Add() fails otherwise.
        try:
            v.Delete()
        except Exception:
            pass

        v_type = _V_TYPE[validation_type]
        op = _V_OPERATOR.get(operator_key, _V_OPERATOR["between"])

        if validation_type == "list":
            # Prefer `formula` (range reference like "=Sheet2!A:A") over listValues.
            source = formula if formula else (",".join(str(x) for x in (list_values or [])))
            if not source:
                return {"status": "error", "message": "addValidation list requires 'formula' or 'listValues'."}
            v.Add(Type=v_type, AlertStyle=_ALERT_STOP, Formula1=source)
            v.InCellDropdown = True

        elif validation_type in ("wholeNumber", "decimal", "textLength"):
            f1 = str(minv) if minv is not None else "0"
            if maxv is not None:
                v.Add(Type=v_type, AlertStyle=_ALERT_STOP, Operator=op, Formula1=f1, Formula2=str(maxv))
            else:
                v.Add(Type=v_type, AlertStyle=_ALERT_STOP, Operator=op, Formula1=f1)

        elif validation_type == "date":
            # Only pass formulae that were supplied — date rules tolerate missing bounds.
            kwargs: dict[str, Any] = {"Type": v_type, "AlertStyle": _ALERT_STOP, "Operator": op}
            if minv is not None:
                kwargs["Formula1"] = str(minv)
            if maxv is not None:
                kwargs["Formula2"] = str(maxv)
            # Validation.Add requires at least Formula1; supply an empty string fallback.
            if "Formula1" not in kwargs:
                kwargs["Formula1"] = ""
            v.Add(**kwargs)

        elif validation_type == "custom":
            f = formula or "=TRUE"
            v.Add(Type=v_type, AlertStyle=_ALERT_STOP, Formula1=f)

        # Error alert configuration.
        v.ShowError = bool(show_error_alert)
        if show_error_alert:
            v.ErrorTitle = "Invalid Input"
            v.ErrorMessage = error_message

    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Validation failed: {exc}",
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"Applied {validation_type} validation to {rng.address}.",
        "outputs": {"range": rng.address, "validationType": validation_type},
    }


registry.register("addValidation", handler, mutates=False, affects_formatting=False)
