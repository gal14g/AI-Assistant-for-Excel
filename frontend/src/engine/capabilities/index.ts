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

export { registry } from "../capabilityRegistry";
