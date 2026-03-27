"""
Pydantic models for the execution plan schema.

These mirror the TypeScript types in frontend/src/engine/types.ts.
The LLM planner produces plans conforming to these models.
The validator checks them before they are sent to the frontend.
"""

from __future__ import annotations

from enum import Enum
from typing import Optional, Union
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
    returnColumns: list[int]
    matchType: str = "exact"
    outputRange: str
    preferFormula: Optional[bool] = True


class GroupSumParams(BaseModel):
    dataRange: str
    groupByColumn: int
    sumColumn: int
    outputRange: str
    preferFormula: Optional[bool] = True
    includeHeaders: Optional[bool] = None


class CreateTableParams(BaseModel):
    range: str
    tableName: str
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
    sortFields: list[SortField]
    hasHeaders: Optional[bool] = True


class PivotValue(BaseModel):
    field: str
    summarizeBy: str = "sum"
    displayName: Optional[str] = None


class CreatePivotParams(BaseModel):
    sourceRange: str
    destinationRange: str
    pivotName: str
    rows: list[str]
    columns: Optional[list[str]] = None
    values: list[PivotValue]
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
}
