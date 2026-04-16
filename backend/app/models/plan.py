"""
Pydantic models for the execution plan schema.

These mirror the TypeScript types in frontend/src/engine/types.ts.
The LLM planner produces plans conforming to these models.
The validator checks them before they are sent to the frontend.
"""

from __future__ import annotations

from enum import Enum
from typing import Literal, Optional, Union
from pydantic import BaseModel, Field


class StepAction(str, Enum):
    readRange = "readRange"
    writeValues = "writeValues"
    writeFormula = "writeFormula"
    matchRecords = "matchRecords"
    groupSum = "groupSum"
    createTable = "createTable"
    applyFilter = "applyFilter"
    sortRange = "sortRange"
    createPivot = "createPivot"
    createChart = "createChart"
    addConditionalFormat = "addConditionalFormat"
    cleanupText = "cleanupText"
    removeDuplicates = "removeDuplicates"
    freezePanes = "freezePanes"
    findReplace = "findReplace"
    addValidation = "addValidation"
    addSheet = "addSheet"
    renameSheet = "renameSheet"
    deleteSheet = "deleteSheet"
    copySheet = "copySheet"
    protectSheet = "protectSheet"
    autoFitColumns = "autoFitColumns"
    mergeCells = "mergeCells"
    setNumberFormat = "setNumberFormat"
    insertDeleteRows = "insertDeleteRows"
    addSparkline = "addSparkline"
    formatCells = "formatCells"
    clearRange = "clearRange"
    hideShow = "hideShow"
    addComment = "addComment"
    addHyperlink = "addHyperlink"
    groupRows = "groupRows"
    setRowColSize = "setRowColSize"
    copyPasteRange = "copyPasteRange"
    pageLayout = "pageLayout"
    insertPicture = "insertPicture"
    insertShape = "insertShape"
    insertTextBox = "insertTextBox"
    addSlicer = "addSlicer"
    splitColumn = "splitColumn"
    unpivot = "unpivot"
    crossTabulate = "crossTabulate"
    bulkFormula = "bulkFormula"
    compareSheets = "compareSheets"
    consolidateRanges = "consolidateRanges"
    extractPattern = "extractPattern"
    categorize = "categorize"
    fillBlanks = "fillBlanks"
    subtotals = "subtotals"
    transpose = "transpose"
    namedRange = "namedRange"
    # --- New actions (batch 2) ---
    fuzzyMatch = "fuzzyMatch"
    deleteRowsByCondition = "deleteRowsByCondition"
    splitByGroup = "splitByGroup"
    lookupAll = "lookupAll"
    regexReplace = "regexReplace"
    coerceDataType = "coerceDataType"
    normalizeDates = "normalizeDates"
    deduplicateAdvanced = "deduplicateAdvanced"
    joinSheets = "joinSheets"
    frequencyDistribution = "frequencyDistribution"
    runningTotal = "runningTotal"
    rankColumn = "rankColumn"
    topN = "topN"
    percentOfTotal = "percentOfTotal"
    growthRate = "growthRate"
    consolidateAllSheets = "consolidateAllSheets"
    cloneSheetStructure = "cloneSheetStructure"
    addReportHeader = "addReportHeader"
    alternatingRowFormat = "alternatingRowFormat"
    quickFormat = "quickFormat"
    refreshPivot = "refreshPivot"
    pivotCalculatedField = "pivotCalculatedField"
    addDropdownControl = "addDropdownControl"
    conditionalFormula = "conditionalFormula"
    spillFormula = "spillFormula"
    # --- New actions (batch 3) ---
    lateralSpreadDuplicates = "lateralSpreadDuplicates"
    extractMatchedToNewRow = "extractMatchedToNewRow"
    reorderRows = "reorderRows"
    fillSeries = "fillSeries"
    insertDeleteColumns = "insertDeleteColumns"
    setSheetDirection = "setSheetDirection"
    tabColor = "tabColor"
    sheetPosition = "sheetPosition"
    autoFitRows = "autoFitRows"
    calculationMode = "calculationMode"
    highlightDuplicates = "highlightDuplicates"
    concatRows = "concatRows"
    insertBlankRows = "insertBlankRows"
    # --- New actions (batch 4 — analytical primitives) ---
    tieredFormula = "tieredFormula"
    histogram = "histogram"
    forecast = "forecast"
    aging = "aging"
    pareto = "pareto"


# --- Step parameter models ---


class ReadRangeParams(BaseModel):
    range: str
    includeHeaders: Optional[bool] = None


class WriteValuesParams(BaseModel):
    range: str
    values: list[list[Union[str, int, float, bool, None]]]
    valuesOnly: Optional[bool] = True


class WriteFormulaParams(BaseModel):
    cell: str
    formula: str
    fillDown: Optional[int] = None


class MatchRecordsParams(BaseModel):
    lookupRange: str
    sourceRange: str
    returnColumns: Optional[list[int]] = None  # defaults to [1] (first source column)
    matchType: Literal["exact", "contains", "approximate"] = "exact"
    outputRange: str
    preferFormula: Optional[bool] = True
    # When set, write this constant string for matched rows (forces composite key matching)
    writeValue: Optional[str] = None


class GroupSumParams(BaseModel):
    dataRange: str
    groupByColumn: int
    sumColumn: int
    outputRange: str
    preferFormula: Optional[bool] = True
    includeHeaders: Optional[bool] = None


class CreateTableParams(BaseModel):
    range: str
    tableName: Optional[str] = None  # auto-generated if omitted
    hasHeaders: Optional[bool] = True
    style: Optional[str] = None


class FilterCriteria(BaseModel):
    filterOn: Literal["values", "topItems", "custom"]
    values: Optional[list[str]] = None
    operator: Optional[str] = None
    value: Optional[Union[str, int, float]] = None


class ApplyFilterParams(BaseModel):
    tableNameOrRange: str
    columnIndex: int
    criteria: FilterCriteria


class SortField(BaseModel):
    columnIndex: int
    ascending: Optional[bool] = True


class SortRangeParams(BaseModel):
    range: str
    sortFields: Optional[list[SortField]] = None  # defaults to first column ascending
    hasHeaders: Optional[bool] = True


class PivotValue(BaseModel):
    field: str
    summarizeBy: str = "sum"
    displayName: Optional[str] = None


class CreatePivotParams(BaseModel):
    sourceRange: str
    destinationRange: Optional[str] = None   # new sheet created if omitted
    pivotName: Optional[str] = None          # auto-generated if omitted
    rows: Optional[list[str]] = None         # auto-detected from headers if omitted
    columns: Optional[list[str]] = None
    values: Optional[list[PivotValue]] = None  # auto-detected from headers if omitted
    filters: Optional[list[str]] = None


class CreateChartParams(BaseModel):
    dataRange: str
    chartType: Literal[
        "columnClustered", "columnStacked", "bar", "line", "pie", "area", "scatter", "combo",
    ]
    title: Optional[str] = None
    sheetName: Optional[str] = None
    position: Optional[dict] = None
    seriesNames: Optional[list[str]] = None


class AddConditionalFormatParams(BaseModel):
    range: str
    ruleType: Literal["cellValue", "formula", "colorScale", "dataBar", "iconSet", "text"]
    operator: Optional[str] = None
    values: Optional[list[Union[str, int, float]]] = None
    format: Optional[dict] = None
    text: Optional[str] = None
    formula: Optional[str] = None  # For formula-based conditional formats


class CleanupTextParams(BaseModel):
    range: str
    operations: list[str]
    outputRange: Optional[str] = None


class RemoveDuplicatesParams(BaseModel):
    range: str
    columnIndexes: Optional[list[int]] = None


class FreezePanesParams(BaseModel):
    cell: str
    sheetName: Optional[str] = None


class FindReplaceParams(BaseModel):
    range: Optional[str] = None
    sheetName: Optional[str] = None
    find: str
    replace: str
    matchCase: Optional[bool] = False
    matchEntireCell: Optional[bool] = False


class AddValidationParams(BaseModel):
    range: str
    validationType: Literal["list", "wholeNumber", "decimal", "date", "textLength", "custom"]
    listValues: Optional[list[str]] = None
    operator: Optional[str] = None
    min: Optional[Union[int, float, str]] = None
    max: Optional[Union[int, float, str]] = None
    formula: Optional[str] = None
    showErrorAlert: Optional[bool] = True
    errorMessage: Optional[str] = None


class AddSheetParams(BaseModel):
    sheetName: str

class RenameSheetParams(BaseModel):
    sheetName: str
    newName: str

class DeleteSheetParams(BaseModel):
    sheetName: str

class CopySheetParams(BaseModel):
    sheetName: str
    newName: Optional[str] = None

class ProtectSheetParams(BaseModel):
    sheetName: str
    password: Optional[str] = None

class AutoFitColumnsParams(BaseModel):
    range: Optional[str] = None
    sheetName: Optional[str] = None

class MergeCellsParams(BaseModel):
    range: str
    mergeType: Optional[Literal["merge", "mergeAcross", "mergeAllCells"]] = "merge"

class SetNumberFormatParams(BaseModel):
    range: str
    format: str  # e.g. "#,##0.00", "0%", "dd/mm/yyyy"


class InsertDeleteRowsParams(BaseModel):
    range: str
    shiftDirection: Literal["down", "up", "right", "left"]


class AddSparklineParams(BaseModel):
    dataRange: str
    locationRange: str
    sparklineType: Optional[Literal["line", "column", "winLoss"]] = "line"
    color: Optional[str] = None


class BordersParams(BaseModel):
    style: str = "thin"
    color: Optional[str] = None
    edges: Optional[list[str]] = None


class FormatCellsParams(BaseModel):
    range: str
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    strikethrough: Optional[bool] = None
    fontSize: Optional[int] = None
    fontFamily: Optional[str] = None
    fontColor: Optional[str] = None
    fillColor: Optional[str] = None
    horizontalAlignment: Optional[str] = None
    verticalAlignment: Optional[str] = None
    wrapText: Optional[bool] = None
    borders: Optional[BordersParams] = None


class ClearRangeParams(BaseModel):
    range: str
    clearType: Literal["contents", "formats", "all"] = "contents"


class HideShowParams(BaseModel):
    target: Literal["sheet", "rows", "columns"]
    rangeOrName: str
    hide: bool = True


class AddCommentParams(BaseModel):
    cell: str
    content: str
    author: Optional[str] = None


class AddHyperlinkParams(BaseModel):
    cell: str
    url: str
    displayText: Optional[str] = None


class GroupRowsParams(BaseModel):
    range: str
    operation: Literal["group", "ungroup"] = "group"


class SetRowColSizeParams(BaseModel):
    range: str
    dimension: Literal["rowHeight", "columnWidth"]
    size: float


class CopyPasteRangeParams(BaseModel):
    sourceRange: str
    destinationRange: str
    pasteType: Optional[Literal["all", "values", "formats", "formulas"]] = "all"


class PageLayoutMargins(BaseModel):
    top: Optional[float] = None
    bottom: Optional[float] = None
    left: Optional[float] = None
    right: Optional[float] = None
    header: Optional[float] = None
    footer: Optional[float] = None


class PageLayoutParams(BaseModel):
    sheetName: Optional[str] = None
    margins: Optional[PageLayoutMargins] = None
    orientation: Optional[Literal["portrait", "landscape"]] = None
    paperSize: Optional[str] = None
    printArea: Optional[str] = None
    showGridlines: Optional[bool] = None
    printGridlines: Optional[bool] = None


class InsertPictureParams(BaseModel):
    sheetName: Optional[str] = None
    imageBase64: str
    left: Optional[float] = None
    top: Optional[float] = None
    width: Optional[float] = None
    height: Optional[float] = None
    altText: Optional[str] = None


class InsertShapeParams(BaseModel):
    sheetName: Optional[str] = None
    shapeType: str
    left: float
    top: float
    width: float
    height: float
    fillColor: Optional[str] = None
    lineColor: Optional[str] = None
    lineWeight: Optional[float] = None
    textContent: Optional[str] = None


class InsertTextBoxParams(BaseModel):
    sheetName: Optional[str] = None
    text: str
    left: float
    top: float
    width: float
    height: float
    fontSize: Optional[float] = None
    fontFamily: Optional[str] = None
    fontColor: Optional[str] = None
    fillColor: Optional[str] = None
    horizontalAlignment: Optional[str] = None


class AddSlicerParams(BaseModel):
    sheetName: Optional[str] = None
    sourceType: Literal["pivotTable", "table"]
    sourceName: str
    sourceField: str
    left: Optional[float] = None
    top: Optional[float] = None
    width: Optional[float] = None
    height: Optional[float] = None


class SplitColumnParams(BaseModel):
    sourceRange: str
    delimiter: str
    outputStartColumn: str
    outputHeaders: Optional[list[str]] = None
    parts: Optional[int] = None


class UnpivotParams(BaseModel):
    sourceRange: str
    idColumns: int
    outputRange: str
    variableColumnName: Optional[str] = None
    valueColumnName: Optional[str] = None


class CrossTabulateParams(BaseModel):
    sourceRange: str
    rowField: int
    columnField: int
    valueField: int
    aggregation: Literal["count", "sum", "average"]
    outputRange: str


class BulkFormulaParams(BaseModel):
    formula: str
    outputRange: str
    dataRange: str
    hasHeaders: Optional[bool] = True


class CompareSheetsParams(BaseModel):
    rangeA: str
    rangeB: str
    outputRange: Optional[str] = None
    highlightDiffs: Optional[bool] = False
    highlightColor: Optional[str] = None


class ConsolidateRangesParams(BaseModel):
    sourceRanges: list[str]
    outputRange: str
    direction: Optional[Literal["vertical", "horizontal"]] = "vertical"
    addSourceLabel: Optional[bool] = False
    deduplicate: Optional[bool] = False


class ExtractPatternParams(BaseModel):
    sourceRange: str
    pattern: str  # built-in name or custom regex
    outputRange: str
    allMatches: Optional[bool] = False


class CategorizeRule(BaseModel):
    operator: Literal[
        "contains", "equals", "startsWith", "endsWith", "greaterThan", "lessThan", "regex",
    ]
    value: Union[str, int, float]
    label: str


class CategorizeParams(BaseModel):
    sourceRange: str
    outputRange: str
    rules: list[CategorizeRule]
    defaultValue: Optional[str] = None


class FillBlanksParams(BaseModel):
    range: str
    fillMode: Optional[Literal["down", "up", "constant"]] = "down"
    constantValue: Optional[Union[str, int, float]] = None


class SubtotalsParams(BaseModel):
    dataRange: str
    groupByColumn: int
    subtotalColumns: list[int]
    aggregation: Optional[Literal["sum", "count", "average"]] = "sum"
    subtotalLabel: Optional[str] = None


class TransposeParams(BaseModel):
    sourceRange: str
    outputRange: str
    copyFormatting: Optional[bool] = False


class NamedRangeParams(BaseModel):
    operation: Literal["create", "update", "delete"]
    name: str
    range: Optional[str] = None
    sheetName: Optional[str] = None


# --- New param models (batch 2) ---


class FuzzyMatchParams(BaseModel):
    lookupRange: str
    sourceRange: str
    outputRange: str
    threshold: Optional[float] = 0.7
    writeValue: Optional[str] = None
    returnBestMatch: Optional[bool] = True


class DeleteRowsByConditionParams(BaseModel):
    range: str
    column: int  # 1-based
    condition: Literal["blank", "notBlank", "equals", "notEquals", "contains", "greaterThan", "lessThan"]
    value: Optional[Union[str, int, float]] = None
    hasHeaders: Optional[bool] = True


class SplitByGroupParams(BaseModel):
    dataRange: str
    groupByColumn: int  # 1-based
    keepHeaders: Optional[bool] = True


class LookupAllParams(BaseModel):
    lookupRange: str
    sourceRange: str
    returnColumn: int  # 1-based
    outputRange: str
    delimiter: Optional[str] = ", "
    matchType: Optional[Literal["exact", "contains"]] = "exact"


class RegexReplaceParams(BaseModel):
    range: str
    pattern: str
    replacement: str
    flags: Optional[str] = "gi"


class CoerceDataTypeParams(BaseModel):
    range: str
    targetType: Literal["number", "text", "date"]
    dateFormat: Optional[str] = None
    locale: Optional[str] = None


class NormalizeDatesParams(BaseModel):
    range: str
    outputFormat: str  # e.g. "yyyy-mm-dd", "dd/mm/yyyy"
    inputFormat: Optional[str] = None


class DeduplicateAdvancedParams(BaseModel):
    range: str
    keyColumns: list[int]  # 1-based
    keepStrategy: Literal["first", "last", "mostComplete", "newest"] = "first"
    dateColumn: Optional[int] = None  # 1-based, for "newest"
    hasHeaders: Optional[bool] = True


class JoinSheetsParams(BaseModel):
    leftRange: str
    rightRange: str
    leftKeyColumn: int  # 1-based
    rightKeyColumn: int  # 1-based
    joinType: Literal["inner", "left", "right", "full"] = "inner"
    outputRange: str


class FrequencyDistributionParams(BaseModel):
    sourceRange: str
    outputRange: str
    sortBy: Optional[Literal["value", "frequency"]] = "frequency"
    ascending: Optional[bool] = False
    includePercent: Optional[bool] = True


class RunningTotalParams(BaseModel):
    sourceRange: str
    outputRange: str
    hasHeaders: Optional[bool] = True


class RankColumnParams(BaseModel):
    sourceRange: str
    outputRange: str
    order: Optional[Literal["descending", "ascending"]] = "descending"
    hasHeaders: Optional[bool] = True


class TopNParams(BaseModel):
    dataRange: str
    valueColumn: int  # 1-based
    n: int
    position: Optional[Literal["top", "bottom"]] = "top"
    outputRange: str
    hasHeaders: Optional[bool] = True


class PercentOfTotalParams(BaseModel):
    sourceRange: str
    outputRange: str
    hasHeaders: Optional[bool] = True
    formatAsPercent: Optional[bool] = True


class GrowthRateParams(BaseModel):
    sourceRange: str
    outputRange: str
    hasHeaders: Optional[bool] = True
    formatAsPercent: Optional[bool] = True


class ConsolidateAllSheetsParams(BaseModel):
    outputSheetName: Optional[str] = "Combined"
    hasHeaders: Optional[bool] = True
    excludeSheets: Optional[list[str]] = None


class CloneSheetStructureParams(BaseModel):
    sourceSheet: str
    newSheetName: str


class AddReportHeaderParams(BaseModel):
    title: str
    sheetName: Optional[str] = None
    range: Optional[str] = None
    fontSize: Optional[int] = 16
    fillColor: Optional[str] = "#4472C4"
    fontColor: Optional[str] = "#FFFFFF"
    bold: Optional[bool] = True


class AlternatingRowFormatParams(BaseModel):
    range: str
    evenColor: Optional[str] = "#F2F2F2"
    oddColor: Optional[str] = "#FFFFFF"
    hasHeaders: Optional[bool] = True


class QuickFormatParams(BaseModel):
    range: str
    freezeHeader: Optional[bool] = True
    addFilters: Optional[bool] = True
    autoFit: Optional[bool] = True
    zebraStripe: Optional[bool] = False
    headerColor: Optional[str] = "#4472C4"
    headerFontColor: Optional[str] = "#FFFFFF"


class RefreshPivotParams(BaseModel):
    pivotName: Optional[str] = None
    sheetName: Optional[str] = None


class PivotCalculatedFieldParams(BaseModel):
    pivotName: str
    sheetName: Optional[str] = None
    fieldName: str
    formula: str


class AddDropdownControlParams(BaseModel):
    cell: str
    listSource: str
    promptMessage: Optional[str] = None
    sheetName: Optional[str] = None


class ConditionalFormulaParams(BaseModel):
    range: str
    conditionColumn: int  # 1-based
    condition: Literal["blank", "notBlank", "equals", "notEquals", "contains", "greaterThan", "lessThan"]
    conditionValue: Optional[Union[str, int, float]] = None
    trueFormula: str  # template with {row} placeholder
    falseFormula: str  # template with {row} placeholder
    outputRange: str
    hasHeaders: Optional[bool] = True


class SpillFormulaParams(BaseModel):
    cell: str
    formula: str
    sheetName: Optional[str] = None


class LateralSpreadDuplicatesParams(BaseModel):
    """Lift every non-first-occurrence row of `keyColumnIndex` out of vertical
    position and paste it horizontally into columns on `direction` side of the
    first-occurrence row. Produces a 'duplicate sidecar' layout used to review
    repeated entries side-by-side (e.g. interview follow-ups, order revisions)."""
    sourceRange: str               # Full data range (e.g. "Sheet5!A1:G100")
    keyColumnIndex: int = Field(..., ge=0)   # 0-based column with the dedupe key
    hasHeaders: Optional[bool] = True
    direction: Optional[Literal["left", "right"]] = "left"
    removeOriginalDuplicates: Optional[bool] = True
    sheetName: Optional[str] = None


class ExtractMatchedToNewRowParams(BaseModel):
    """When row[keyColumnIndexA] == row[keyColumnIndexB] (same value in two
    designated columns of the SAME row), extract the values at
    `extractColumnIndexes` into a new row inserted below. The shared key value
    is duplicated into the column-A position of the new row so the extracted
    record is identifiable. Original positions of extracted cells are blanked
    on the matched row. Useful for normalizing side-by-side comparison data
    (interview first/second round, email primary/secondary, etc.)."""
    sourceRange: str
    keyColumnIndexA: int = Field(..., ge=0)
    keyColumnIndexB: int = Field(..., ge=0)
    extractColumnIndexes: list[int] = Field(..., min_length=1)
    hasHeaders: Optional[bool] = True
    caseSensitive: Optional[bool] = False
    sheetName: Optional[str] = None


class ReorderRowsParams(BaseModel):
    """Reorder rows in a range. Mode 'moveMatching' moves rows whose
    `conditionColumn` satisfies `condition`/`conditionValue` to the
    destination (top/bottom of range or just after a specified row index).
    Mode 'reverse' flips the row order. Mode 'clusterByKey' groups rows
    sharing the same value in `conditionColumn` together, preserving first-
    appearance order of the key."""
    range: str
    mode: Literal["moveMatching", "reverse", "clusterByKey"]
    conditionColumn: Optional[int] = Field(None, ge=0)   # 0-based
    condition: Optional[Literal["equals", "notEquals", "contains", "notContains", "blank", "notBlank", "greaterThan", "lessThan"]] = None
    conditionValue: Optional[Union[str, int, float]] = None
    destination: Optional[Literal["top", "bottom"]] = "top"
    hasHeaders: Optional[bool] = True
    sheetName: Optional[str] = None


class FillSeriesParams(BaseModel):
    """Write a generated series into a range. seriesType: 'number' increments
    by step (default 1); 'date' steps by `dateUnit` days/weeks/months; 'weekday'
    skips Sat/Sun; 'repeatPattern' cycles through `pattern` values."""
    range: str
    seriesType: Literal["number", "date", "weekday", "repeatPattern"]
    start: Optional[Union[str, int, float]] = 1
    step: Optional[Union[int, float]] = 1
    pattern: Optional[list[Union[str, int, float, bool]]] = None  # for repeatPattern
    dateUnit: Optional[Literal["day", "week", "month", "year"]] = "day"
    count: Optional[int] = None     # if omitted, fill the entire range
    horizontal: Optional[bool] = False
    sheetName: Optional[str] = None


class InsertDeleteColumnsParams(BaseModel):
    """Insert or delete columns. range is a column-letter range like 'C:E'
    (or a cell-address range from which the column span is inferred)."""
    range: str
    action: Literal["insert", "delete"]
    shiftDirection: Optional[Literal["left", "right"]] = "right"
    sheetName: Optional[str] = None


class SetSheetDirectionParams(BaseModel):
    """Request sheet right-to-left / left-to-right display. NOTE: Office.js
    has no API for this; the add-in handler returns a descriptive warning so
    the user can toggle manually via View > Sheet Right-to-Left. In MCP mode
    (xlwings desktop bridge), this will actually set the COM property."""
    direction: Literal["rtl", "ltr"]
    sheetName: Optional[str] = None


class TabColorParams(BaseModel):
    """Set the color of a worksheet's tab. `color` is a hex string (e.g. '#FF0000')
    or 'none' to clear."""
    color: str
    sheetName: Optional[str] = None


class SheetPositionParams(BaseModel):
    """Move a sheet to position `position` (0-based) in the tab order."""
    position: int = Field(..., ge=0)
    sheetName: Optional[str] = None


class AutoFitRowsParams(BaseModel):
    """Auto-fit row heights for the given range, or the used range if omitted."""
    range: Optional[str] = None
    sheetName: Optional[str] = None


class CalculationModeParams(BaseModel):
    """Set workbook calculation mode. 'manual' disables auto-recalc; 'automatic'
    re-enables it; 'automaticExceptTables' auto-recalcs everything except
    data-tables."""
    mode: Literal["manual", "automatic", "automaticExceptTables"]


class HighlightDuplicatesParams(BaseModel):
    """Add a conditional-formatting rule that highlights duplicate values in
    the given range. Combines addConditionalFormat + COUNTIF in one step."""
    range: str
    fillColor: Optional[str] = "#FFCCCC"
    fontColor: Optional[str] = "#C50F1F"
    sheetName: Optional[str] = None


class ConcatRowsParams(BaseModel):
    """Concatenate the cells of each row in `sourceRange` into a single cell
    in `outputColumn`, joined by `separator`. Uses TEXTJOIN under the hood so
    the output formulas stay live as source values change."""
    sourceRange: str
    outputColumn: str              # e.g. "G" or "Sheet1!G"
    separator: Optional[str] = ", "
    ignoreBlanks: Optional[bool] = True
    hasHeaders: Optional[bool] = True
    sheetName: Optional[str] = None


class InsertBlankRowsParams(BaseModel):
    """Insert blank rows either at a set of explicit row numbers or at a fixed
    interval (every Nth row within a range)."""
    sheetName: Optional[str] = None
    positions: Optional[list[int]] = None  # 1-based row numbers to insert before
    every: Optional[int] = Field(None, ge=1)            # e.g. every 5 rows
    range: Optional[str] = None            # confines "every" mode to this range
    count: Optional[int] = 1               # how many blank rows per insertion


# ── Analytical primitives (batch 4) ────────────────────────────────────────


class TieredFormulaTier(BaseModel):
    """One tier: values ≥ `threshold` use `value`. In 'tax' mode, `value`
    is a rate applied to the slice of income above `threshold`."""
    threshold: Union[int, float]
    value: Union[int, float]


class TieredFormulaParams(BaseModel):
    """Apply tiered logic (tax brackets, grading bands, commission tiers) via a
    generated IFS / nested-IF formula. Mode 'lookup' picks the tier's value
    whose threshold ≤ source cell. Mode 'tax' computes cumulative tier tax —
    `value` for each tier is a rate applied to the slice of the source
    between that tier's threshold and the next tier's threshold."""
    sourceRange: str                # single-column range of inputs
    outputRange: str                # single-column range for results
    tiers: list[TieredFormulaTier] = Field(..., min_length=1)
    mode: Optional[Literal["lookup", "tax"]] = "lookup"
    defaultValue: Optional[Union[int, float]] = 0
    hasHeaders: Optional[bool] = True


class HistogramParams(BaseModel):
    """Build a histogram of `dataRange`. Either supply explicit `bins` edges
    or request `binCount` automatic bins (Sturges' rule). Emits the FREQUENCY
    formula + a column chart (unless `includeChart=false`)."""
    dataRange: str                  # single-column numeric data
    outputRange: str                # top-left for the bin-count output (2 cols wide: bin, count)
    bins: Optional[list[Union[int, float]]] = None
    binCount: Optional[int] = Field(None, ge=1)
    includeChart: Optional[bool] = True
    chartType: Optional[Literal["columnClustered", "barClustered"]] = "columnClustered"
    hasHeaders: Optional[bool] = True
    sheetName: Optional[str] = None


class ForecastParams(BaseModel):
    """Project future values from a time series. `sourceRange` is two columns:
    dates in col 1, values in col 2. Emits either a FORECAST.LINEAR or
    FORECAST.ETS formula for each of the next `periods` date steps. Writes a
    continuation table + optional line chart."""
    sourceRange: str
    outputRange: str                # top-left for the projection table (2 cols: date, forecast)
    periods: int = Field(..., ge=1)
    method: Optional[Literal["linear", "ets"]] = "linear"
    includeChart: Optional[bool] = True
    hasHeaders: Optional[bool] = True
    sheetName: Optional[str] = None


class AgingParams(BaseModel):
    """Bucket dates into age buckets (e.g. 0-30, 31-60, 61-90, 90+). Each
    bucket's upper-bound day count is supplied via `buckets` (sorted asc);
    open-ended last bucket is labelled `{lastBucket}+`. Writes a formula
    column next to the source."""
    dateColumn: str                 # single column of dates
    outputColumn: str               # letter like "G" or "Sheet1!G"
    buckets: Optional[list[int]] = None     # default [30, 60, 90]
    referenceDate: Optional[str] = None     # e.g. "01/04/2026" — default TODAY()
    hasHeaders: Optional[bool] = True
    sheetName: Optional[str] = None


class ParetoParams(BaseModel):
    """Pareto (80/20) analysis: sort by value descending, write the sorted
    labels + values + cumulative-percentage columns, and produce a combo
    chart (column for value, line for cumulative %). Useful for identifying
    the top drivers of any aggregated measure."""
    dataRange: str                  # 2 cols: label + value
    outputRange: str                # top-left of the output block (label, value, cum %)
    includeChart: Optional[bool] = True
    hasHeaders: Optional[bool] = True
    sheetName: Optional[str] = None


# --- Plan step ---


class PlanStep(BaseModel):
    id: str = Field(..., min_length=1)
    description: str = Field(..., min_length=1)
    # Display-only translation of `description` for non-English users.
    # The canonical `description` is always English; UIs prefer this when set.
    # See the LANGUAGE RULE in services/chat_service.py system prompt.
    descriptionLocalized: Optional[str] = None
    action: StepAction
    params: dict  # Validated further based on action
    dependsOn: Optional[list[str]] = None


# --- Execution plan ---


class ExecutionPlan(BaseModel):
    planId: str = Field(..., min_length=1)
    createdAt: str
    userRequest: str
    summary: str
    # Display-only translation of `summary`. See PlanStep.descriptionLocalized.
    summaryLocalized: Optional[str] = None
    steps: list[PlanStep] = Field(..., min_length=1)
    preserveFormatting: bool = True
    confidence: float = Field(ge=0, le=1)
    warnings: Optional[list[str]] = None


# Map action to param model for detailed validation.
# Every StepAction must have an entry here so params are always checked.
ACTION_PARAM_MODELS: dict[StepAction, type[BaseModel]] = {
    StepAction.readRange:            ReadRangeParams,
    StepAction.writeValues:          WriteValuesParams,
    StepAction.writeFormula:         WriteFormulaParams,
    StepAction.matchRecords:         MatchRecordsParams,
    StepAction.groupSum:             GroupSumParams,
    StepAction.createTable:          CreateTableParams,
    StepAction.applyFilter:          ApplyFilterParams,
    StepAction.sortRange:            SortRangeParams,
    StepAction.createPivot:          CreatePivotParams,
    StepAction.createChart:          CreateChartParams,
    StepAction.addConditionalFormat: AddConditionalFormatParams,
    StepAction.cleanupText:          CleanupTextParams,
    StepAction.removeDuplicates:     RemoveDuplicatesParams,
    StepAction.freezePanes:          FreezePanesParams,
    StepAction.findReplace:          FindReplaceParams,
    StepAction.addValidation:        AddValidationParams,
    StepAction.addSheet:             AddSheetParams,
    StepAction.renameSheet:          RenameSheetParams,
    StepAction.deleteSheet:          DeleteSheetParams,
    StepAction.copySheet:            CopySheetParams,
    StepAction.protectSheet:         ProtectSheetParams,
    StepAction.autoFitColumns:       AutoFitColumnsParams,
    StepAction.mergeCells:           MergeCellsParams,
    StepAction.setNumberFormat:      SetNumberFormatParams,
    StepAction.insertDeleteRows:     InsertDeleteRowsParams,
    StepAction.addSparkline:         AddSparklineParams,
    StepAction.formatCells:          FormatCellsParams,
    StepAction.clearRange:           ClearRangeParams,
    StepAction.hideShow:             HideShowParams,
    StepAction.addComment:           AddCommentParams,
    StepAction.addHyperlink:         AddHyperlinkParams,
    StepAction.groupRows:            GroupRowsParams,
    StepAction.setRowColSize:        SetRowColSizeParams,
    StepAction.copyPasteRange:       CopyPasteRangeParams,
    StepAction.pageLayout:           PageLayoutParams,
    StepAction.insertPicture:        InsertPictureParams,
    StepAction.insertShape:          InsertShapeParams,
    StepAction.insertTextBox:        InsertTextBoxParams,
    StepAction.addSlicer:            AddSlicerParams,
    StepAction.splitColumn:          SplitColumnParams,
    StepAction.unpivot:              UnpivotParams,
    StepAction.crossTabulate:        CrossTabulateParams,
    StepAction.bulkFormula:          BulkFormulaParams,
    StepAction.compareSheets:        CompareSheetsParams,
    StepAction.consolidateRanges:    ConsolidateRangesParams,
    StepAction.extractPattern:       ExtractPatternParams,
    StepAction.categorize:           CategorizeParams,
    StepAction.fillBlanks:           FillBlanksParams,
    StepAction.subtotals:            SubtotalsParams,
    StepAction.transpose:            TransposeParams,
    StepAction.namedRange:           NamedRangeParams,
    # --- New actions (batch 2) ---
    StepAction.fuzzyMatch:              FuzzyMatchParams,
    StepAction.deleteRowsByCondition:   DeleteRowsByConditionParams,
    StepAction.splitByGroup:            SplitByGroupParams,
    StepAction.lookupAll:               LookupAllParams,
    StepAction.regexReplace:            RegexReplaceParams,
    StepAction.coerceDataType:          CoerceDataTypeParams,
    StepAction.normalizeDates:          NormalizeDatesParams,
    StepAction.deduplicateAdvanced:     DeduplicateAdvancedParams,
    StepAction.joinSheets:              JoinSheetsParams,
    StepAction.frequencyDistribution:   FrequencyDistributionParams,
    StepAction.runningTotal:            RunningTotalParams,
    StepAction.rankColumn:              RankColumnParams,
    StepAction.topN:                    TopNParams,
    StepAction.percentOfTotal:          PercentOfTotalParams,
    StepAction.growthRate:              GrowthRateParams,
    StepAction.consolidateAllSheets:    ConsolidateAllSheetsParams,
    StepAction.cloneSheetStructure:     CloneSheetStructureParams,
    StepAction.addReportHeader:         AddReportHeaderParams,
    StepAction.alternatingRowFormat:    AlternatingRowFormatParams,
    StepAction.quickFormat:             QuickFormatParams,
    StepAction.refreshPivot:            RefreshPivotParams,
    StepAction.pivotCalculatedField:    PivotCalculatedFieldParams,
    StepAction.addDropdownControl:      AddDropdownControlParams,
    StepAction.conditionalFormula:      ConditionalFormulaParams,
    StepAction.spillFormula:            SpillFormulaParams,
    # --- New actions (batch 3) ---
    StepAction.lateralSpreadDuplicates: LateralSpreadDuplicatesParams,
    StepAction.extractMatchedToNewRow:  ExtractMatchedToNewRowParams,
    StepAction.reorderRows:             ReorderRowsParams,
    StepAction.fillSeries:              FillSeriesParams,
    StepAction.insertDeleteColumns:     InsertDeleteColumnsParams,
    StepAction.setSheetDirection:       SetSheetDirectionParams,
    StepAction.tabColor:                TabColorParams,
    StepAction.sheetPosition:           SheetPositionParams,
    StepAction.autoFitRows:             AutoFitRowsParams,
    StepAction.calculationMode:         CalculationModeParams,
    StepAction.highlightDuplicates:     HighlightDuplicatesParams,
    StepAction.concatRows:              ConcatRowsParams,
    StepAction.insertBlankRows:         InsertBlankRowsParams,
    # --- batch 4 (analytical primitives) ---
    StepAction.tieredFormula:           TieredFormulaParams,
    StepAction.histogram:               HistogramParams,
    StepAction.forecast:                ForecastParams,
    StepAction.aging:                   AgingParams,
    StepAction.pareto:                  ParetoParams,
}
