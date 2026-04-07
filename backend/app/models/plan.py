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
    filterOn: str  # "values" | "topItems" | "custom"
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
    chartType: str
    title: Optional[str] = None
    sheetName: Optional[str] = None
    position: Optional[dict] = None
    seriesNames: Optional[list[str]] = None


class AddConditionalFormatParams(BaseModel):
    range: str
    ruleType: str
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
    validationType: str
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
    mergeType: Optional[str] = "merge"  # "merge" | "mergeAcross" | "mergeAllCells"

class SetNumberFormatParams(BaseModel):
    range: str
    format: str  # e.g. "#,##0.00", "0%", "dd/mm/yyyy"


class InsertDeleteRowsParams(BaseModel):
    range: str
    shiftDirection: str  # "down" | "up" | "right" | "left"


class AddSparklineParams(BaseModel):
    dataRange: str
    locationRange: str
    sparklineType: Optional[str] = "line"  # "line" | "column" | "winLoss"
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
    clearType: str = "contents"


class HideShowParams(BaseModel):
    target: str
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
    operation: str = "group"


class SetRowColSizeParams(BaseModel):
    range: str
    dimension: str
    size: float


class CopyPasteRangeParams(BaseModel):
    sourceRange: str
    destinationRange: str
    pasteType: Optional[str] = "all"


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
    orientation: Optional[str] = None  # "portrait" | "landscape"
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
    sourceType: str  # "pivotTable" | "table"
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
    aggregation: str  # "count" | "sum" | "average"
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
    direction: Optional[str] = "vertical"  # "vertical" | "horizontal"
    addSourceLabel: Optional[bool] = False
    deduplicate: Optional[bool] = False


class ExtractPatternParams(BaseModel):
    sourceRange: str
    pattern: str  # built-in name or custom regex
    outputRange: str
    allMatches: Optional[bool] = False


class CategorizeRule(BaseModel):
    operator: str  # "contains" | "equals" | "startsWith" | "endsWith" | "greaterThan" | "lessThan" | "regex"
    value: Union[str, int, float]
    label: str


class CategorizeParams(BaseModel):
    sourceRange: str
    outputRange: str
    rules: list[CategorizeRule]
    defaultValue: Optional[str] = None


class FillBlanksParams(BaseModel):
    range: str
    fillMode: Optional[str] = "down"  # "down" | "up" | "constant"
    constantValue: Optional[Union[str, int, float]] = None


class SubtotalsParams(BaseModel):
    dataRange: str
    groupByColumn: int
    subtotalColumns: list[int]
    aggregation: Optional[str] = "sum"  # "sum" | "count" | "average"
    subtotalLabel: Optional[str] = None


class TransposeParams(BaseModel):
    sourceRange: str
    outputRange: str
    copyFormatting: Optional[bool] = False


class NamedRangeParams(BaseModel):
    operation: str  # "create" | "update" | "delete"
    name: str
    range: Optional[str] = None
    sheetName: Optional[str] = None


# --- Plan step ---


class PlanStep(BaseModel):
    id: str = Field(..., min_length=1)
    description: str = Field(..., min_length=1)
    action: StepAction
    params: dict  # Validated further based on action
    dependsOn: Optional[list[str]] = None


# --- Execution plan ---


class ExecutionPlan(BaseModel):
    planId: str = Field(..., min_length=1)
    createdAt: str
    userRequest: str
    summary: str
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
}
