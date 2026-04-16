/**
 * Capability Index
 *
 * Importing this module registers ALL capabilities with the registry.
 * This file must be imported once at app startup.
 */

import "./readRange";
import "./writeValues";
import "./writeFormula";
import "./matchRecords";
import "./groupSum";
import "./createTable";
import "./applyFilter";
import "./sortRange";
import "./createPivot";
import "./createChart";
import "./conditionalFormat";
import "./cleanupText";
import "./removeDuplicates";
import "./freezePanes";
import "./findReplace";
import "./sheetOps";
import "./validation";
import "./autoFitColumns";
import "./mergeCells";
import "./setNumberFormat";
import "./insertDeleteRows";
import "./addSparkline";
import "./formatCells";
import "./clearRange";
import "./hideShow";
import "./addComment";
import "./addHyperlink";
import "./groupRows";
import "./setRowColSize";
import "./copyPasteRange";
import "./pageLayout";
import "./insertPicture";
import "./insertShape";
import "./insertTextBox";
import "./addSlicer";
import "./splitColumn";
import "./unpivot";
import "./crossTabulate";
import "./bulkFormula";
import "./compareSheets";
import "./consolidateRanges";
import "./extractPattern";
import "./categorize";
import "./fillBlanks";
import "./subtotals";
import "./transpose";
import "./namedRange";
import "./fuzzyMatch";
import "./deleteRowsByCondition";
import "./splitByGroup";
import "./lookupAll";
import "./regexReplace";
import "./runningTotal";
import "./rankColumn";
import "./topN";
import "./percentOfTotal";
import "./growthRate";
import "./coerceDataType";
import "./normalizeDates";
import "./deduplicateAdvanced";
import "./joinSheets";
import "./frequencyDistribution";
import "./addDropdownControl";
import "./spillFormula";
import "./consolidateAllSheets";
import "./cloneSheetStructure";
import "./addReportHeader";
import "./alternatingRowFormat";
import "./quickFormat";
import "./refreshPivot";
import "./pivotCalculatedField";
import "./conditionalFormula";
import "./lateralSpreadDuplicates";
// --- batch 3 (row reshape + sheet ops + series generation) ---
import "./extractMatchedToNewRow";
import "./reorderRows";
import "./fillSeries";
import "./insertDeleteColumns";
import "./setSheetDirection";
import "./tabColor";
import "./sheetPosition";
import "./autoFitRows";
import "./calculationMode";
import "./highlightDuplicates";
import "./concatRows";
import "./insertBlankRows";
// --- batch 4 (analytical primitives) ---
import "./tieredFormula";
import "./histogram";
import "./forecast";
import "./aging";
import "./pareto";

export { registry } from "../capabilityRegistry";
