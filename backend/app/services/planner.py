"""
Capability catalog + LLM response-parsing utilities.

CAPABILITY_DESCRIPTIONS is the source-of-truth docstring used when building
the chat system prompt (services/chat_service.py) and when indexing
capabilities for vector retrieval (services/capability_store.py).

extract_json() handles stripping markdown fences, trailing commas, and
malformed output from LLM responses before parsing.
"""

from __future__ import annotations

import json

CAPABILITY_DESCRIPTIONS = {
    "readRange": "Read values from a cell range. Params: range (string), includeHeaders (bool, optional)",
    "writeValues": "Write a 2D array of values to a range. ONLY writes values, never formatting. Params: range (string), values (2D array)",
    "writeFormula": "Write a formula to a cell, optionally fill down. PREFER this over writeValues when a native Excel formula can express the operation. Params: cell (string), formula (string starting with =), fillDown (int, optional)",
    "matchRecords": "Lookup/match records between ranges using XLOOKUP/VLOOKUP or composite key matching. Params: lookupRange, sourceRange, returnColumns (array of 1-based ints), matchType ('exact'|'contains'|'approximate'), outputRange, preferFormula (bool, default true). IMPORTANT: when the user says 'contained in' / 'exists in' / 'is in' / 'מוכל', ALWAYS set matchType='contains' — this uses substring matching so partial values match. 'exact' requires identical values. SPECIAL: when the user wants to write a constant string (like 'pass' or 'v') to a column for matched rows, set writeValue='pass' instead of returnColumns — this triggers composite key matching. When writeValue is set, preferFormula is ignored. NEVER use writeValues to write match results row-by-row.",
    "groupSum": "Sum values grouped by a column using SUMIF or computed aggregation. Params: dataRange, groupByColumn (1-based int), sumColumn (1-based int), outputRange, preferFormula (bool, default true), includeHeaders (bool, default true — set false if data has no header row)",
    "createTable": "Convert a range into an Excel Table. Params: range, tableName, hasHeaders (bool), style (optional string)",
    "applyFilter": "Apply filters to a table or range. Params: tableNameOrRange, columnIndex (0-based), criteria {filterOn: 'values'|'custom', values: string[] (for filterOn='values'), operator: 'greaterThan'|'lessThan'|'equals'|'contains' (for filterOn='custom'), value: string (for filterOn='custom')}",
    "sortRange": "Sort a range by columns. Params: range, sortFields [{columnIndex (0-based), ascending}], hasHeaders (bool)",
    "createPivot": "Create a PivotTable. Only sourceRange is required — everything else is auto-detected. rows/values fields accept EITHER header names ('Department') OR column range addresses ('Sheet2!A:A') — the handler resolves range addresses to header names automatically. Params: sourceRange (required), rows (optional list of field names or column refs), values (optional list of {field, summarizeBy} where field is a name or ref and summarizeBy is 'sum'|'count'|'average'|'max'|'min'), columns (optional), filters (optional), destinationRange (optional — new sheet created if omitted), pivotName (optional)",
    "createChart": "Create a chart. Params: dataRange, chartType ('columnClustered'|'columnStacked'|'bar'|'line'|'pie'|'area'|'scatter'|'combo'), title (optional), position (optional)",
    "addConditionalFormat": "Apply conditional formatting. Params: range, ruleType ('cellValue'|'colorScale'|'dataBar'|'iconSet'|'text'|'formula'), operator ('greaterThan'|'greaterThanOrEqualTo'|'lessThan'|'lessThanOrEqualTo'|'between'|'notBetween'|'equalTo'|'notEqualTo') — used with ruleType='cellValue', values, format {fillColor, fontColor, bold}, text (string — for ruleType='text': highlights cells containing this substring), formula (for ruleType='formula': Excel formula string e.g. '=$D2=\"\"' to highlight blank rows, '=$B2>$C2' for cross-column compare)",
    "cleanupText": "Clean up text values. Params: range, operations ['trim'|'lowercase'|'uppercase'|'properCase'|'removeNonPrintable'|'normalizeWhitespace'], outputRange (optional)",
    "removeDuplicates": "Remove duplicate rows. Params: range, columnIndexes (0-based array, optional)",
    "freezePanes": "Freeze rows/columns at a cell. Params: cell (string), sheetName (optional)",
    "findReplace": "Find and replace text. Params: find, replace, range (optional), matchCase (bool), matchEntireCell (bool). Supports date-aware matching (finds dates in different formats). If replace starts with '=' it writes a formula instead of a value.",
    "addValidation": "Add data validation. Params: range, validationType ('list'|'wholeNumber'|'decimal'|'date'|'textLength'|'custom'), listValues (comma list for static dropdown), formula ('=Sheet2!A:A' for dynamic range dropdown OR custom formula for validationType='custom'), operator ('between'|'notBetween'|'equalTo'|'notEqualTo'|'greaterThan'|'greaterThanOrEqualTo'|'lessThan'|'lessThanOrEqualTo'), min, max",
    "addSheet": "Add a new worksheet. Params: sheetName",
    "renameSheet": "Rename a worksheet. Params: sheetName, newName",
    "deleteSheet": "Delete a worksheet. Params: sheetName",
    "copySheet": "Copy a worksheet. Params: sheetName, newName (optional)",
    "protectSheet":    "Protect a worksheet. Params: sheetName, password (optional)",
    "autoFitColumns":    "Auto-fit column widths to their content. Params: range (optional — omit to fit all used columns), sheetName (optional)",
    "mergeCells":        "Merge cells in a range. Params: range (string), mergeType ('merge'=everything into one cell, 'mergeAcross'=each row separately)",
    "setNumberFormat":   "Apply a number format to a range. Params: range (string), format (e.g. '#,##0.00', '0%', 'dd/mm/yyyy', '$#,##0.00', 'General')",
    "insertDeleteRows":  "Insert or delete rows/columns. Params: range (determines which rows/columns and how many), shiftDirection ('down'=insert rows above, 'up'=delete rows, 'right'=insert columns left, 'left'=delete columns)",
    "addSparkline":      "Add sparkline mini-charts inside cells — ideal for dashboards showing trends. Params: dataRange (source data, one row per sparkline), locationRange (cells where sparklines appear), sparklineType ('line'|'column'|'winLoss', default 'line'), color (optional hex)",
    "formatCells":       "Format cell appearance — font, colors, borders, alignment. Params: range, bold (bool), italic (bool), underline (bool), strikethrough (bool), fontSize (int), fontFamily (string e.g. 'Calibri'), fontColor (hex), fillColor (hex), horizontalAlignment ('left'|'center'|'right'|'justify'), verticalAlignment ('top'|'middle'|'bottom'), wrapText (bool), borders ({style: 'thin'|'medium'|'thick'|'dashed'|'dotted'|'double'|'none', color: hex, edges: ['top'|'bottom'|'left'|'right'|'all'|'outside'|'inside']}). All params except range are optional — only set what you need to change.",
    "clearRange":        "Clear a range's contents, formatting, or both. Params: range, clearType ('contents'=values+formulas, 'formats'=only formatting, 'all'=everything)",
    "hideShow":          "Hide or unhide rows, columns, or entire sheets. Params: target ('rows'|'columns'|'sheet'), rangeOrName (row range e.g. '2:5', column range e.g. 'A:C', or sheet name), hide (bool: true=hide, false=unhide)",
    "addComment":        "Add a comment/note to a cell. Params: cell (string), content (string), author (optional string)",
    "addHyperlink":      "Insert a hyperlink in a cell. Params: cell (string), url (string), displayText (optional — defaults to url)",
    "groupRows":         "Group or ungroup rows/columns for outline collapsing. Params: range (row range e.g. '3:8' or column range e.g. 'B:E'), operation ('group'|'ungroup')",
    "setRowColSize":     "Set row height or column width manually. Params: range (row range e.g. '1:1' or column range e.g. 'A:C'), dimension ('rowHeight'|'columnWidth'), size (number — points for rows, characters for columns)",
    "copyPasteRange":    "Copy a range and paste to another location. Params: sourceRange, destinationRange, pasteType ('all'|'values'|'formats'|'formulas', default 'all')",
    "pageLayout": "Set page layout: margins, orientation, paper size, print area, gridline visibility",
    "insertPicture": "Insert an image (base64) into a worksheet at a given position and size",
    "insertShape": "Insert a geometric shape (rectangle, oval, arrow, star, etc.) with fill, outline, and optional text",
    "insertTextBox": "Insert a text box with styled content at a given position",
    "addSlicer": "Add a slicer control for filtering a PivotTable or Table by a specific field",
    "splitColumn": "Split a text column into multiple columns by a delimiter (e.g. 'John Smith' → 'John' | 'Smith'). Params: sourceRange (single-column range), delimiter (string — use empty string for fixed-width), outputStartColumn (column letter where first output goes, e.g. 'B'), outputHeaders (list of strings, optional), parts (int, optional — how many pieces to split into).",
    "unpivot": "Reshape wide data into tall format (melt). The left idColumns stay; remaining columns collapse into variable + value columns. Params: sourceRange (range including headers), idColumns (int — number of left columns to keep as IDs), outputRange, variableColumnName (optional, default 'Attribute'), valueColumnName (optional, default 'Value').",
    "crossTabulate": "Build a contingency matrix / cross-tab from raw data. Params: sourceRange (raw data including headers), rowField (1-based column index for matrix rows), columnField (1-based column index for matrix columns), valueField (1-based column index for the aggregated value), aggregation ('count'|'sum'|'average'), outputRange.",
    "bulkFormula": "Fill a formula template across a whole output column, sized to the dataRange. Use this when the user says 'add this formula to the whole column'. Params: formula (template for first data row, e.g. '=A2*B2'), outputRange (target column range), dataRange (range that defines how many rows to fill), hasHeaders (bool, default true — skips row 1).",
    "compareSheets": "Compare two ranges cell-by-cell and optionally highlight differences or write a diff report. Params: rangeA, rangeB, outputRange (optional — where to write the diff report; new sheet created if omitted), highlightDiffs (bool, default false — color diff cells in rangeA), highlightColor (hex, optional).",
    "consolidateRanges": "Combine multiple ranges into one — vertical stack (append rows) or horizontal join (side-by-side). Params: sourceRanges (list of range strings), outputRange, direction ('vertical'|'horizontal', default 'vertical'), addSourceLabel (bool — prepends source range as a column), deduplicate (bool, default false).",
    "extractPattern": "Extract a pattern from each cell using a regex or a named built-in. Params: sourceRange (source column), pattern ('email'|'phone'|'url'|'date'|'number'|'currency' or a custom regex string), outputRange, allMatches (bool, default false — if true, join all matches per cell).",
    "categorize": "Classify each row in a column into labels by applying rules (first match wins). Params: sourceRange (source column), outputRange, rules (list of {operator: 'contains'|'equals'|'startsWith'|'endsWith'|'greaterThan'|'lessThan'|'regex', value, label}), defaultValue (optional string for no-rule-matches).",
    "fillBlanks": "Forward-fill or back-fill blank cells — ideal for cleaning merged-cell exports. Params: range, fillMode ('down'=carry value from above (default) | 'up'=carry value from below | 'constant'), constantValue (required when fillMode='constant').",
    "subtotals": "Insert subtotal rows after each group in sorted data (Excel's Data→Subtotals). Params: dataRange (including headers), groupByColumn (1-based int), subtotalColumns (list of 1-based ints to aggregate), aggregation ('sum'|'count'|'average', default 'sum'), subtotalLabel (optional, default 'Total'). Sorts by groupByColumn first.",
    "transpose": "Swap rows and columns of a range (Excel's Paste Special → Transpose). Params: sourceRange, outputRange — outputRange dimensions must be the inverse of sourceRange, copyFormatting (bool, default false — values only).",
    "namedRange": "Create, update, or delete a named range. Params: operation ('create'|'update'|'delete'), name (string), range (required for create/update), sheetName (optional — if provided, range is sheet-scoped; omit for workbook-scoped).",
}


def _clean_json_text(text: str) -> str:
    """
    Fix common LLM JSON generation issues before parsing:
    - trailing commas before } or ]
    - stray control characters
    """
    import re
    # Remove trailing commas before closing braces/brackets
    text = re.sub(r",\s*([}\]])", r"\1", text)
    # Remove control characters except standard whitespace
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    return text


def extract_json(text: str) -> dict:
    """
    Extract a JSON object from LLM response text.

    Handles:
    - Bare JSON (starts with {)
    - Markdown code fences (```json ... ```)
    - Trailing commas and minor formatting issues
    - Falls back to json-repair for deeply malformed output
    """
    import re

    text = text.strip()

    # Strip markdown fences first
    if "```json" in text:
        m = re.search(r"```json\s*(.*?)```", text, re.DOTALL)
        if m:
            text = m.group(1).strip()
    elif "```" in text:
        m = re.search(r"```\s*(.*?)```", text, re.DOTALL)
        if m:
            candidate = m.group(1).strip()
            if candidate.startswith("{"):
                text = candidate

    # Extract the outermost JSON object if surrounded by prose
    if not text.startswith("{"):
        try:
            first_brace = text.index("{")
            last_brace = text.rindex("}")
            text = text[first_brace : last_brace + 1]
        except ValueError:
            pass

    # First try: standard parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Second try: clean trailing commas etc. then parse
    cleaned = _clean_json_text(text)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    # Third try: json-repair (handles deeply malformed LLM output)
    try:
        from json_repair import repair_json
        repaired = repair_json(cleaned, return_objects=True)
        if isinstance(repaired, dict):
            return repaired
        return json.loads(repair_json(cleaned))
    except Exception:
        pass

    raise ValueError(f"Could not extract valid JSON from LLM response. First 200 chars: {text[:200]}")
