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
    # --- New capabilities (batch 2) ---
    "fuzzyMatch": "Fuzzy string matching between two columns using Levenshtein distance. Finds approximate matches even when values aren't identical (e.g. 'IBM' vs 'I.B.M.'). Params: lookupRange (column to match from), sourceRange (column to match against), outputRange, threshold (0-1 similarity, default 0.7), writeValue (optional constant to write for matches instead of matched value), returnBestMatch (bool, default true — write the best matching source value).",
    "deleteRowsByCondition": "Delete rows where a column meets a condition. Params: range (data range including headers), column (1-based int), condition ('blank'|'notBlank'|'equals'|'notEquals'|'contains'|'greaterThan'|'lessThan'), value (for equals/notEquals/contains/greaterThan/lessThan), hasHeaders (bool, default true).",
    "splitByGroup": "Split data into separate sheets by unique values in a column — opposite of consolidate. Params: dataRange (range including headers), groupByColumn (1-based int), keepHeaders (bool, default true — include header in each new sheet).",
    "lookupAll": "Find ALL matching rows between two ranges (not just the first match like XLOOKUP). Writes all matched values joined by delimiter. Params: lookupRange, sourceRange, returnColumn (1-based int), outputRange, delimiter (default ', '), matchType ('exact'|'contains', default 'exact').",
    "regexReplace": "Regex find-and-replace across a range. Supports capture groups ($1, $2). Params: range, pattern (regex string), replacement (can use $1, $2 capture groups), flags (default 'gi' — global + case-insensitive).",
    "coerceDataType": "Convert column values from one type to another (text→number, text→date, number→text). Fixes 'stored as text' issues. Params: range, targetType ('number'|'text'|'date'), dateFormat (optional), locale (optional).",
    "normalizeDates": "Standardize mixed date formats in a column to one consistent format. Handles Excel serial numbers, dd/mm/yyyy, mm/dd/yyyy, yyyy-mm-dd, etc. Params: range, outputFormat (e.g. 'yyyy-mm-dd', 'dd/mm/yyyy'), inputFormat (optional hint).",
    "deduplicateAdvanced": "Remove duplicates with control over which row to keep. Params: range, keyColumns (list of 1-based ints), keepStrategy ('first'|'last'|'mostComplete'|'newest'), dateColumn (1-based int, required for 'newest'), hasHeaders (bool, default true).",
    "joinSheets": "SQL-style join between two ranges on key columns. Params: leftRange, rightRange, leftKeyColumn (1-based int), rightKeyColumn (1-based int), joinType ('inner'|'left'|'right'|'full'), outputRange.",
    "frequencyDistribution": "Count occurrences of each unique value and write a frequency table. Params: sourceRange, outputRange, sortBy ('value'|'frequency', default 'frequency'), ascending (bool, default false), includePercent (bool, default true).",
    "runningTotal": "Write cumulative sum formulas (running total). Params: sourceRange (value column), outputRange, hasHeaders (bool, default true).",
    "rankColumn": "Write RANK formulas for values in a column. Params: sourceRange, outputRange, order ('descending'|'ascending', default 'descending'), hasHeaders (bool, default true).",
    "topN": "Extract the top N or bottom N rows sorted by a value column. Params: dataRange, valueColumn (1-based int), n (int), position ('top'|'bottom', default 'top'), outputRange, hasHeaders (bool, default true).",
    "percentOfTotal": "Write percentage-of-total formulas. Params: sourceRange (value column), outputRange, hasHeaders (bool, default true), formatAsPercent (bool, default true — applies 0.0% number format).",
    "growthRate": "Calculate period-over-period growth rate formulas. Params: sourceRange (value column), outputRange, hasHeaders (bool, default true), formatAsPercent (bool, default true).",
    "consolidateAllSheets": "Merge data from ALL worksheets into one combined sheet. Params: outputSheetName (default 'Combined'), hasHeaders (bool, default true — only include headers from first sheet), excludeSheets (optional list of sheet names to skip).",
    "cloneSheetStructure": "Copy a sheet's structure (headers + formatting + column widths) without data. Params: sourceSheet (name), newSheetName.",
    "addReportHeader": "Insert a formatted report title row above data — inserts row, merges across, applies styling. Params: title (string), sheetName (optional), range (optional — insert above this range), fontSize (default 16), fillColor (default '#4472C4'), fontColor (default '#FFFFFF'), bold (default true).",
    "alternatingRowFormat": "Apply zebra-stripe formatting to alternate rows. Params: range, evenColor (hex, default '#F2F2F2'), oddColor (hex, default '#FFFFFF'), hasHeaders (bool, default true — skip header row).",
    "quickFormat": "Apply a combination of common formatting in one step: freeze header + add filters + auto-fit + optional zebra-stripe. Params: range, freezeHeader (default true), addFilters (default true), autoFit (default true), zebraStripe (default false), headerColor (default '#4472C4'), headerFontColor (default '#FFFFFF').",
    "refreshPivot": "Refresh a PivotTable or all PivotTables on a sheet. Params: pivotName (optional), sheetName (optional — if neither, refreshes all on active sheet).",
    "pivotCalculatedField": "Add a calculated field to an existing PivotTable. Params: pivotName, sheetName (optional), fieldName, formula.",
    "addDropdownControl": "Create a dropdown (data validation list) in a cell. Params: cell, listSource (comma-separated values like 'A,B,C' or a range reference like 'Sheet2!A:A'), promptMessage (optional), sheetName (optional).",
    "conditionalFormula": "Write IF-based formulas that apply different logic depending on a condition. Params: range (data range), conditionColumn (1-based int), condition ('blank'|'notBlank'|'equals'|'notEquals'|'contains'|'greaterThan'|'lessThan'), conditionValue (for equals/contains/etc.), trueFormula (template with {row} placeholder e.g. '=B{row}*1.1'), falseFormula (template with {row}), outputRange, hasHeaders (default true).",
    "spillFormula": "Write a dynamic array formula (FILTER, SORT, UNIQUE, SEQUENCE, RANDARRAY) that spills automatically. Params: cell (single cell), formula (e.g. '=FILTER(A:C,B:B>100)'), sheetName (optional). IMPORTANT: the formula MUST reference a bounded range (e.g. 'A2:A200'), NEVER a full-column reference like 'A:A' — that spills to 1,048,576 cells and will be rejected.",
    "lateralSpreadDuplicates": "Build a duplicate-sidecar layout: for each non-first-occurrence of a key column, lift that row's entire data out of vertical position and paste it horizontally next to the first-occurrence row (on the left or right). Removes the original duplicate rows. Use this when the user asks to 'show duplicates side-by-side', 'move duplicates next to their first occurrence', or similar lateral-join-on-self requests. Params: sourceRange (full data range including headers, e.g. 'Sheet1!A1:G100'), keyColumnIndex (0-based column used as the dedupe key), hasHeaders (default true), direction ('left' default, or 'right'), removeOriginalDuplicates (default true), sheetName (optional). Hebrew directional words about row repositioning — משמאל/מימין/למעלה/למטה — are LIST-ORDER hints, not physical axes; usually 'להיות משמאל/לפני' means 'before' in reading order.",
    "extractMatchedToNewRow": "When row[keyColumnIndexA] == row[keyColumnIndexB] (two designated columns of the SAME row hold matching values), extract the values at extractColumnIndexes into a NEW row inserted immediately below the matched row. The shared key value is copied into the column-A position of the new row so the extracted record is identifiable on its own. Values at the extracted positions on the matched row are blanked. Useful for normalizing side-by-side comparison data (interview first/second round columns, primary/secondary email, request/fulfilled pairs). Params: sourceRange, keyColumnIndexA (0-based), keyColumnIndexB (0-based), extractColumnIndexes (0-based list, usually the 'B-side' columns whose values should be extracted), hasHeaders (default true), caseSensitive (default false), sheetName (optional).",
    "reorderRows": "Reorder rows in a range without copying to another sheet. mode='moveMatching' relocates rows whose conditionColumn satisfies condition/conditionValue to 'top' or 'bottom' of the range. mode='reverse' flips the row order. mode='clusterByKey' groups rows sharing the same value in conditionColumn together (preserving first-appearance order of the key). Params: range, mode, conditionColumn (0-based, required for moveMatching/clusterByKey), condition ('equals'|'notEquals'|'contains'|'notContains'|'blank'|'notBlank'|'greaterThan'|'lessThan'), conditionValue, destination ('top' default or 'bottom' for moveMatching), hasHeaders (default true), sheetName (optional).",
    "fillSeries": "Write a generated series into a range. seriesType='number' increments by step (default 1). seriesType='date' steps by dateUnit ('day'|'week'|'month'|'year', default 'day'). seriesType='weekday' is dates skipping Sat/Sun. seriesType='repeatPattern' cycles through `pattern` values. Params: range, seriesType, start (default 1 or today for dates), step, pattern (for repeatPattern), dateUnit, count (optional — default: fill the entire range), horizontal (default false), sheetName (optional). Use this instead of writeValues with a hand-built 2D array when the user asks for '1 to 100 in A', 'dates for next 60 days', 'numbering', 'weekly schedule', etc.",
    "insertDeleteColumns": "Insert or delete columns — column-axis pair of insertDeleteRows. range is a column-letter range like 'C:E' or a cell-address range whose column span is used. Params: range, action ('insert' or 'delete'), shiftDirection ('left' or 'right', default 'right'), sheetName (optional).",
    "setSheetDirection": "Set sheet direction to right-to-left or left-to-right. IMPORTANT: Office.js does NOT expose a worksheet RTL property; in add-in mode this handler returns a descriptive warning instructing the user to toggle manually via View > Sheet Right-to-Left. In MCP mode (desktop xlwings bridge) this will actually set the COM property. Use when the user asks for RTL/LTR sheet flip (Hebrew/Arabic users). Params: direction ('rtl' or 'ltr'), sheetName (optional).",
    "tabColor": "Set the color of a worksheet's tab. Common for color-coding sheets. Params: color (hex string like '#FF0000' or 'none' to clear), sheetName (optional).",
    "sheetPosition": "Move a sheet to a specific position in the tab order. Params: position (0-based — 0 = leftmost), sheetName (optional).",
    "autoFitRows": "Auto-fit row heights — row-axis pair of autoFitColumns. Params: range (optional — omit to auto-fit all used rows on the active sheet), sheetName (optional).",
    "calculationMode": "Set workbook calculation mode. Useful before large bulk-writes to avoid recalc thrashing. Params: mode ('manual' | 'automatic' | 'automaticExceptTables'). Remember to set back to 'automatic' afterwards.",
    "highlightDuplicates": "Add a conditional-formatting rule that highlights duplicates in a range. One-step equivalent of addConditionalFormat with a COUNTIF formula. Params: range, fillColor (default '#FFCCCC'), fontColor (default '#C50F1F'), sheetName (optional). Use when the user says 'highlight duplicates in column X' or 'color duplicate values'.",
    "concatRows": "Concatenate the cells of each row in sourceRange into a single cell in outputColumn, joined by separator. Uses TEXTJOIN so outputs stay live. Params: sourceRange, outputColumn (letter like 'G' or 'Sheet1!G'), separator (default ', '), ignoreBlanks (default true), hasHeaders (default true), sheetName (optional).",
    "insertBlankRows": "Insert blank rows either at explicit row numbers or every Nth row within a range. Common for print layouts / visual separation. Params: sheetName (optional), positions (1-based row numbers to insert before — use for explicit positions), every (integer >= 1 — use for regular interval), range (required with 'every' to scope the interval), count (how many blank rows per insertion, default 1). Provide EITHER positions OR every+range.",
    "tieredFormula": "Apply tiered logic (tax brackets, grading bands, commission tiers, discount thresholds) via a generated IFS formula. mode='lookup' picks a tier's value when the source cell is ≥ that tier's threshold (first match wins if tiers are sorted descending). mode='tax' computes cumulative tier tax — `value` for each tier is a rate applied to the slice of the source between that tier's threshold and the next tier's threshold. Params: sourceRange (single column), outputRange, tiers (list of {threshold, value} — sorted ascending by threshold), mode ('lookup' default or 'tax'), defaultValue (default 0), hasHeaders (default true). Use this instead of writing a 10-deep nested =IF() by hand.",
    "histogram": "Build a histogram of `dataRange`. Either supply explicit `bins` edges or request `binCount` auto bins (Sturges' rule). Emits a FREQUENCY formula + a column chart (unless includeChart=false). Params: dataRange (single-column numeric), outputRange (top-left for 2-col bin/count output), bins (optional list of edge values), binCount (optional int), includeChart (default true), chartType ('columnClustered' default or 'barClustered'), hasHeaders (default true), sheetName (optional).",
    "forecast": "Project future values from a time series. `sourceRange` is two columns: dates in col 1, values in col 2. Emits FORECAST.LINEAR or FORECAST.ETS formulas for each of the next `periods` date steps, writing a continuation table + optional line chart. Params: sourceRange (2-col), outputRange (top-left for the projection table), periods (int >= 1), method ('linear' default or 'ets'), includeChart (default true), hasHeaders (default true), sheetName (optional).",
    "aging": "Bucket dates into aging buckets (AR aging, ticket age, days-since). `buckets` is the upper-bound day count of each bucket (e.g. [30, 60, 90]); the open-ended last bucket is labelled '{lastBucket}+'. Writes a formula column next to the source. Params: dateColumn (single column), outputColumn (letter like 'G' or 'Sheet1!G'), buckets (default [30, 60, 90]), referenceDate (default today), hasHeaders (default true), sheetName (optional).",
    "pareto": "Pareto (80/20) analysis: sort by value descending, write sorted labels + values + cumulative-% columns, and produce a combo chart (column for value + line for cumulative %). Useful for identifying the top drivers of any aggregated measure. Params: dataRange (2-col: label + value), outputRange (top-left of the 3-col output), includeChart (default true), hasHeaders (default true), sheetName (optional).",
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
