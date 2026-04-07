"""
Capability Store – vector-based semantic search for capability matching.

Uses ChromaDB (local, persistent) + sentence-transformers to embed capability
descriptions and example user queries.  At query time, the user's message is
embedded and the top-K most relevant capabilities are returned.

This lets the LLM see only the relevant actions instead of all 34+, resulting
in smaller prompts, faster responses, and more reliable tool selection.
"""

from __future__ import annotations

import functools
import logging

from ..config import settings

logger = logging.getLogger(__name__)

# ── Module state ─────────────────────────────────────────────────────────────

_collection = None  # chromadb.Collection
_ready = False

# ── Example user queries per capability ──────────────────────────────────────
# These are embedded alongside descriptions for better recall.  Users say
# things like "make a chart" — the description alone ("Create a chart.
# Params: dataRange…") wouldn't match as well.

CAPABILITY_EXAMPLES: dict[str, list[str]] = {
    "readRange": [
        "read the data in column A",
        "show me what's in cells A1 to C20",
        "get the values from this range",
    ],
    "writeValues": [
        "write these values to the sheet",
        "paste this data into cells A1",
        "fill in the table with these numbers",
    ],
    "writeFormula": [
        "add a SUM formula",
        "calculate the average in column B",
        "write an IF formula",
        "use VLOOKUP to find values",
        "create a COUNTIF formula",
        "add a dynamic array formula",
    ],
    "matchRecords": [
        "match column A with Sheet2",
        "lookup values between sheets",
        "find matching records",
        "compare two columns across sheets",
        "write 'pass' where rows match",
        "XLOOKUP from one sheet to another",
    ],
    "groupSum": [
        "sum column B grouped by column A",
        "total sales by category",
        "aggregate values by group",
        "SUMIF grouped by department",
    ],
    "createTable": [
        "convert this range to a table",
        "make an Excel table from my data",
        "create a formatted table",
    ],
    "applyFilter": [
        "filter the table to show only values above 100",
        "show only rows where column B is 'Active'",
        "apply a filter to this data",
    ],
    "sortRange": [
        "sort by column B descending",
        "sort the data alphabetically",
        "order rows by date",
    ],
    "createPivot": [
        "create a pivot table",
        "summarize data with a pivot",
        "make a pivot from this range",
        "group and summarize in a pivot table",
    ],
    "createChart": [
        "create a bar chart",
        "make a pie chart showing sales",
        "graph the results",
        "visualize this data as a line chart",
        "add a chart from my data",
    ],
    "addConditionalFormat": [
        "highlight cells above 100 in red",
        "color scale from green to red",
        "add data bars to column B",
        "highlight rows where column D is blank",
        "conditional formatting based on value",
    ],
    "cleanupText": [
        "trim whitespace from column A",
        "convert text to uppercase",
        "clean up the text data",
        "normalize whitespace in cells",
    ],
    "removeDuplicates": [
        "remove duplicate rows",
        "delete duplicates in this range",
        "deduplicate the data",
    ],
    "freezePanes": [
        "freeze the top row",
        "freeze panes at cell B2",
        "lock the header row in place",
    ],
    "findReplace": [
        "find and replace text",
        "replace all occurrences of 'old' with 'new'",
        "search and replace in column A",
        "ctrl+h replace values",
    ],
    "addValidation": [
        "add a dropdown list to column A",
        "validate that cells contain numbers only",
        "add data validation for dates",
        "create a dropdown from a range",
    ],
    "addSheet": [
        "add a new sheet",
        "create a new worksheet called Dashboard",
    ],
    "renameSheet": [
        "rename the sheet to Summary",
        "change the sheet name",
    ],
    "deleteSheet": [
        "delete this sheet",
        "remove the worksheet",
    ],
    "copySheet": [
        "copy this sheet",
        "duplicate the worksheet",
    ],
    "protectSheet": [
        "protect the sheet with a password",
        "lock the worksheet",
    ],
    "autoFitColumns": [
        "auto-fit column widths",
        "resize columns to fit content",
        "adjust column sizes automatically",
    ],
    "mergeCells": [
        "merge cells A1 to D1",
        "combine these cells into one",
        "merge the header row",
    ],
    "setNumberFormat": [
        "format as currency",
        "change number format to percentage",
        "format dates as dd/mm/yyyy",
        "apply number format",
    ],
    "insertDeleteRows": [
        "insert 3 rows above row 5",
        "delete rows 10 through 15",
        "add a new column before column C",
        "remove empty rows",
    ],
    "addSparkline": [
        "add sparklines for each row",
        "add mini charts showing trends",
        "insert sparkline graphs",
    ],
    "formatCells": [
        "make the header bold",
        "change font color to red",
        "add borders to the table",
        "center align the text",
        "change the background color",
        "set font to Arial size 12",
    ],
    "clearRange": [
        "clear the contents of this range",
        "erase everything in column B",
        "clear formatting from these cells",
        "delete all data in the range",
    ],
    "hideShow": [
        "hide column C",
        "unhide rows 5 to 10",
        "hide this sheet",
        "show hidden columns",
    ],
    "addComment": [
        "add a comment to cell A1",
        "insert a note in this cell",
        "annotate cell B5 with a comment",
    ],
    "addHyperlink": [
        "add a link to cell A1",
        "insert a hyperlink",
        "link this cell to a URL",
    ],
    "groupRows": [
        "group rows 3 through 8",
        "collapse these rows",
        "create an outline group",
        "ungroup columns B to E",
    ],
    "setRowColSize": [
        "set row height to 30",
        "make column A wider",
        "change column width to 20",
        "resize row 1",
    ],
    "copyPasteRange": [
        "copy this range and paste it to Sheet2",
        "duplicate these cells to another location",
        "paste values only from A1:B10 to D1",
        "copy formatting from one range to another",
    ],
    "pageLayout": [
        "set the page to landscape orientation",
        "change margins to 1 inch on all sides",
        "set print area to A1:G20",
        "hide gridlines on this sheet",
        "set paper size to A4",
    ],
    "insertPicture": [
        "insert an image into the sheet",
        "add a logo picture at the top left",
        "embed a base64 image at position 100,50",
    ],
    "insertShape": [
        "insert a rectangle shape",
        "add a red arrow pointing right",
        "draw an oval with blue fill",
        "insert a star shape at position 200,100",
    ],
    "insertTextBox": [
        "add a text box with the title 'Summary'",
        "insert a text box at the top of the sheet",
        "create a text box with bold 14pt text",
    ],
    "addSlicer": [
        "add a slicer for the pivot table",
        "create a slicer to filter by Region",
        "add a slicer for the sales table by category",
    ],
    "splitColumn": [
        "split the full name column into first and last name",
        "break this column apart by commas",
        "split email column on the @ sign",
        "separate city and state into two columns",
    ],
    "unpivot": [
        "unpivot the monthly columns into a tall table",
        "melt wide format into long format",
        "reshape from wide to tall",
        "turn these year columns into rows",
    ],
    "crossTabulate": [
        "cross-tabulate region by product",
        "build a contingency table of category and status",
        "count occurrences of A vs B",
        "make a cross-tab matrix",
    ],
    "bulkFormula": [
        "apply this formula to the whole column",
        "fill this formula down for every row",
        "add this formula to all rows in the data",
    ],
    "compareSheets": [
        "compare Sheet1 and Sheet2 and show differences",
        "find cells that differ between these two ranges",
        "diff the old and new versions",
        "highlight what changed between the sheets",
    ],
    "consolidateRanges": [
        "combine these three ranges into one table",
        "stack data from multiple sheets together",
        "merge these ranges vertically",
        "consolidate data from Q1, Q2, Q3, Q4 sheets",
    ],
    "extractPattern": [
        "extract email addresses from this column",
        "pull phone numbers out of the messy text",
        "get the URLs from column B",
        "extract dates from the description column",
    ],
    "categorize": [
        "categorize rows as corporate or personal",
        "tag each row based on these rules",
        "classify customers into buckets",
        "label each amount as small/medium/large",
    ],
    "fillBlanks": [
        "fill empty cells with the value above",
        "forward-fill the merged category column",
        "fill blanks with zero",
        "carry values down to fill empty rows",
    ],
    "subtotals": [
        "add subtotals by department",
        "insert subtotal rows for each category",
        "group with subtotals by region",
        "create Excel subtotals grouped by product",
    ],
    "transpose": [
        "transpose this range",
        "flip rows and columns",
        "swap rows with columns",
        "paste special transpose",
    ],
    "namedRange": [
        "name this range SalesData",
        "create a named range",
        "give this range a name",
        "define a name for these cells",
    ],
    # --- New capabilities (batch 2) ---
    "fuzzyMatch": [
        "fuzzy match company names between two columns",
        "approximate string matching",
        "find similar names even if spelled differently",
        "match IBM with I.B.M.",
    ],
    "deleteRowsByCondition": [
        "delete all rows where column D is blank",
        "remove rows where status equals cancelled",
        "delete empty rows",
        "remove rows containing error",
    ],
    "splitByGroup": [
        "split this sheet into separate sheets by department",
        "create one sheet per category",
        "separate data into tabs by column A values",
        "split by group into different worksheets",
    ],
    "lookupAll": [
        "find all matching records not just the first",
        "VLOOKUP but return all matches",
        "lookup all instances of each value",
        "list all orders for each customer",
    ],
    "regexReplace": [
        "regex replace across this column",
        "use regular expression to clean up text",
        "replace using a pattern with capture groups",
        "extract and reformat phone numbers using regex",
    ],
    "coerceDataType": [
        "convert this column from text to numbers",
        "change stored-as-text to number",
        "convert text dates to date format",
        "fix numbers stored as text",
    ],
    "normalizeDates": [
        "standardize all dates to yyyy-mm-dd",
        "fix mixed date formats in this column",
        "convert all dates to dd/mm/yyyy",
        "normalize date formats",
    ],
    "deduplicateAdvanced": [
        "remove duplicates but keep the most recent row",
        "deduplicate keeping the row with most data",
        "remove duplicates keep last occurrence",
        "deduplicate by name column keeping newest by date",
    ],
    "joinSheets": [
        "join these two sheets like SQL LEFT JOIN",
        "merge two tables by matching ID column",
        "combine data from two sheets by key",
        "inner join sheet1 and sheet2 on column A",
    ],
    "frequencyDistribution": [
        "count how many times each value appears",
        "frequency distribution of this column",
        "create a frequency table",
        "show value counts",
    ],
    "runningTotal": [
        "add a running total column",
        "cumulative sum of sales",
        "running balance",
        "calculate running total",
    ],
    "rankColumn": [
        "rank these values from highest to lowest",
        "add a rank column",
        "rank employees by score",
        "show ranking for each row",
    ],
    "topN": [
        "show me the top 10 by revenue",
        "extract bottom 5 performers",
        "top 20 products by sales",
        "get the 10 lowest values",
    ],
    "percentOfTotal": [
        "calculate percentage of total for each row",
        "what percent does each item contribute",
        "add a percent of total column",
        "show each row as a percentage of the sum",
    ],
    "growthRate": [
        "calculate month over month growth",
        "year over year growth rate",
        "period over period change",
        "growth percentage between rows",
    ],
    "consolidateAllSheets": [
        "merge all sheets into one",
        "combine every worksheet into a single sheet",
        "consolidate all tabs",
        "stack all sheets together",
    ],
    "cloneSheetStructure": [
        "copy the sheet structure without data",
        "create a blank copy with same headers and formatting",
        "duplicate the template but empty",
        "clone sheet layout only",
    ],
    "addReportHeader": [
        "add a title row above the data",
        "insert a report header",
        "create a formatted title at the top",
        "add a merged header row with styling",
    ],
    "alternatingRowFormat": [
        "zebra stripe the rows",
        "alternate row colors",
        "add banded rows",
        "stripe every other row",
    ],
    "quickFormat": [
        "format this table nicely",
        "freeze header add filters and auto-fit",
        "make this look professional",
        "apply standard table formatting",
    ],
    "refreshPivot": [
        "refresh the pivot table",
        "update pivot data",
        "recalculate the pivot",
        "refresh all pivots on this sheet",
    ],
    "pivotCalculatedField": [
        "add a calculated field to the pivot",
        "create a profit margin field in the pivot table",
        "add a computed column to the pivot",
    ],
    "addDropdownControl": [
        "add a dropdown list in this cell",
        "create a dropdown selector",
        "add a filter dropdown",
        "put a dropdown with options A B C",
    ],
    "conditionalFormula": [
        "if column A is blank use formula X otherwise formula Y",
        "apply different formulas based on a condition",
        "conditional calculation based on column value",
        "if status is active multiply by 1.1 otherwise keep same",
    ],
    "spillFormula": [
        "use FILTER to show only rows where column B > 100",
        "write a UNIQUE formula",
        "create a SORT formula that spills",
        "dynamic array formula with FILTER",
    ],
}


def init_store() -> None:
    """
    Initialise the ChromaDB collection and index all capabilities.

    Safe to call multiple times — skips re-indexing if the document count
    matches (idempotent on restarts).
    """
    global _collection, _ready  # noqa: PLW0603

    from .chroma_client import get_chroma_client, get_embedding_fn
    from .planner import CAPABILITY_DESCRIPTIONS

    client = get_chroma_client()
    embedding_fn = get_embedding_fn()

    # Build the full document corpus: description + example queries per action
    ids: list[str] = []
    documents: list[str] = []
    metadatas: list[dict[str, str]] = []

    for action, description in CAPABILITY_DESCRIPTIONS.items():
        # The description itself
        ids.append(f"{action}_desc")
        documents.append(f"{action}: {description}")
        metadatas.append({"action": action})

        # Example user queries
        for i, example in enumerate(CAPABILITY_EXAMPLES.get(action, [])):
            ids.append(f"{action}_ex_{i}")
            documents.append(example)
            metadatas.append({"action": action})

    expected_count = len(ids)

    # Check if we need to re-index
    existing = client.get_or_create_collection(
        name="capabilities",
        embedding_function=embedding_fn,
    )

    if existing.count() == expected_count:
        logger.info(
            "Capability store already indexed (%d docs) — skipping.",
            expected_count,
        )
        _collection = existing
        _ready = True
        return

    # Re-index: delete old collection and create fresh
    client.delete_collection("capabilities")
    _collection = client.create_collection(
        name="capabilities",
        embedding_function=embedding_fn,
    )

    _collection.add(ids=ids, documents=documents, metadatas=metadatas)
    logger.info("Indexed %d capability documents into ChromaDB.", expected_count)
    _ready = True


def is_ready() -> bool:
    """Whether the capability store has been initialised."""
    return _ready


@functools.lru_cache(maxsize=256)
def search_capabilities(query: str, top_k: int | None = None) -> list[str]:
    """
    Return the action names of the top-K most relevant capabilities
    for the given user query.

    Always includes actions referenced by few-shot examples to avoid
    the LLM seeing example actions not in its available list.
    """
    if not _ready or _collection is None:
        # Fallback: return all actions (no filtering)
        from .planner import CAPABILITY_DESCRIPTIONS

        return list(CAPABILITY_DESCRIPTIONS.keys())

    k = top_k or settings.capability_top_k

    # Query more than k to account for deduplication (multiple docs per action)
    results = _collection.query(
        query_texts=[query],
        n_results=min(k * 3, _collection.count()),
    )

    # Deduplicate by action name while preserving relevance order
    seen: dict[str, None] = {}
    if results and results["metadatas"]:
        for meta in results["metadatas"][0]:
            action = meta["action"]
            if action not in seen:
                seen[action] = None

    return list(seen.keys())[:k]
