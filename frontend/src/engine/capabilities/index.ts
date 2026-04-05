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

export { registry } from "../capabilityRegistry";
