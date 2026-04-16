"""
Capability Store — vector-based semantic search for capability matching.

Backed by the `VectorStore` abstraction (`persistence/factory.py`). The
concrete backend is ChromaDB (default) or pgvector, chosen at startup
via `settings.vector_store_url`.

Flow:
  1. `init_store()` embeds every (action description, example query) pair
     into the `capabilities` collection.
  2. `search_capabilities(query)` returns the top-K most relevant action
     names for a user message. The LLM only sees those, not all 70+
     actions — smaller prompt, faster responses, better tool selection.

Re-indexing: `init_store()` is idempotent. If the collection's document
count already matches the expected corpus size, seeding is skipped. A
full re-index is triggered whenever the expected count changes (e.g.
after adding new example queries below or editing an action description).

The full catalogue of example queries per action lives in
`CAPABILITY_EXAMPLES` below — paraphrase-multilingual-MiniLM-L12-v2 covers
Hebrew natively so queries in either language match without a translation
step.
"""

from __future__ import annotations

import functools
import logging

from ..config import settings

logger = logging.getLogger(__name__)

_ready = False
_expected_count = 0

# ── Example user queries per capability ──────────────────────────────────────
# These are embedded alongside descriptions for better recall. Users say
# things like "make a chart" — the description alone ("Create a chart.
# Params: dataRange…") wouldn't match as well.

CAPABILITY_EXAMPLES: dict[str, list[str]] = {
    "readRange": [
        "read the data in column A",
        "show me what's in cells A1 to C20",
        "get the values from this range",
        "תקרא את הנתונים בעמודה A",
        "תראה לי מה יש בטווח הזה",
    ],
    "writeValues": [
        "write these values to the sheet",
        "paste this data into cells A1",
        "fill in the table with these numbers",
        "תכתוב את הערכים האלה לגיליון",
        "תדביק את הנתונים בתאים",
    ],
    "writeFormula": [
        "add a SUM formula",
        "calculate the average in column B",
        "write an IF formula",
        "use VLOOKUP to find values",
        "create a COUNTIF formula",
        "add a dynamic array formula",
        "תוסיף נוסחת סכום",
        "חשב ממוצע בעמודה",
    ],
    "matchRecords": [
        "match column A with Sheet2",
        "lookup values between sheets",
        "find matching records",
        "compare two columns across sheets",
        "write 'pass' where rows match",
        "XLOOKUP from one sheet to another",
        "התאם בין שני גיליונות",
        "חפש ערכים תואמים",
    ],
    "groupSum": [
        "sum column B grouped by column A",
        "total sales by category",
        "aggregate values by group",
        "SUMIF grouped by department",
        "סיכום מכירות לפי קטגוריה",
        "סכום לפי קבוצה",
    ],
    "createTable": [
        "convert this range to a table",
        "make an Excel table from my data",
        "create a formatted table",
        "תהפוך את הטווח לטבלה",
        "תיצור טבלה מהנתונים",
    ],
    "applyFilter": [
        "filter the table to show only values above 100",
        "show only rows where column B is 'Active'",
        "apply a filter to this data",
        "תסנן את הטבלה לערכים מעל 100",
        "תציג רק שורות פעילות",
    ],
    "sortRange": [
        "sort by column B descending",
        "sort the data alphabetically",
        "order rows by date",
        "מיין לפי עמודה",
        "סדר את הנתונים לפי תאריך",
    ],
    "createPivot": [
        "create a pivot table",
        "summarize data with a pivot",
        "make a pivot from this range",
        "group and summarize in a pivot table",
        "תיצור טבלת ציר",
        "סיכום נתונים עם פיבוט",
    ],
    "createChart": [
        "create a bar chart",
        "make a pie chart showing sales",
        "graph the results",
        "visualize this data as a line chart",
        "add a chart from my data",
        "תיצור גרף עמודות",
        "גרף עוגה של מכירות",
    ],
    "addConditionalFormat": [
        "highlight cells above 100 in red",
        "color scale from green to red",
        "add data bars to column B",
        "highlight rows where column D is blank",
        "conditional formatting based on value",
        "עיצוב מותנה לפי ערך",
        "תסמן בצבע תאים מעל 100",
    ],
    "cleanupText": [
        "trim whitespace from column A",
        "convert text to uppercase",
        "clean up the text data",
        "normalize whitespace in cells",
        "נקה רווחים מעמודה A",
        "המר טקסט לאותיות גדולות",
    ],
    "removeDuplicates": [
        "remove duplicate rows",
        "delete duplicates in this range",
        "deduplicate the data",
        "הסר שורות כפולות",
        "מחק כפילויות",
    ],
    "freezePanes": [
        "freeze the top row",
        "freeze panes at cell B2",
        "lock the header row in place",
        "הקפא את השורה העליונה",
        "קבע את הכותרת",
    ],
    "findReplace": [
        "find and replace text",
        "replace all occurrences of 'old' with 'new'",
        "search and replace in column A",
        "ctrl+h replace values",
        "חפש והחלף טקסט",
        "החלף את כל המופעים",
    ],
    "addValidation": [
        "add a dropdown list to column A",
        "validate that cells contain numbers only",
        "add data validation for dates",
        "create a dropdown from a range",
        "תוסיף רשימה נפתחת",
        "אימות נתונים למספרים בלבד",
    ],
    "addSheet": [
        "add a new sheet",
        "create a new worksheet called Dashboard",
        "תיצור גיליון חדש",
        "הוסף גיליון",
    ],
    "renameSheet": [
        "rename the sheet to Summary",
        "change the sheet name",
        "שנה את שם הגיליון",
        "תקרא לגיליון סיכום",
    ],
    "deleteSheet": [
        "delete this sheet",
        "remove the worksheet",
        "מחק את הגיליון הזה",
        "הסר את הגיליון",
    ],
    "copySheet": [
        "copy this sheet",
        "duplicate the worksheet",
        "העתק את הגיליון",
        "שכפל את הגיליון",
    ],
    "protectSheet": [
        "protect the sheet with a password",
        "lock the worksheet",
        "הגן על הגיליון בסיסמה",
        "נעל את הגיליון",
    ],
    "autoFitColumns": [
        "auto-fit column widths",
        "resize columns to fit content",
        "adjust column sizes automatically",
        "התאם רוחב עמודות אוטומטית",
        "שנה גודל עמודות לפי תוכן",
    ],
    "mergeCells": [
        "merge cells A1 to D1",
        "combine these cells into one",
        "merge the header row",
        "מזג תאים",
        "אחד את התאים האלה",
    ],
    "setNumberFormat": [
        "format as currency",
        "change number format to percentage",
        "format dates as dd/mm/yyyy",
        "apply number format",
        "עצב כמטבע",
        "שנה לפורמט אחוזים",
    ],
    "insertDeleteRows": [
        "insert 3 rows above row 5",
        "delete rows 10 through 15",
        "add a new column before column C",
        "remove empty rows",
        "הוסף 3 שורות מעל שורה 5",
        "מחק שורות ריקות",
    ],
    "addSparkline": [
        "add sparklines for each row",
        "add mini charts showing trends",
        "insert sparkline graphs",
        "הוסף גרפי מיני לכל שורה",
        "תוסיף ספארקליינים",
    ],
    "formatCells": [
        "make the header bold",
        "change font color to red",
        "add borders to the table",
        "center align the text",
        "change the background color",
        "set font to Arial size 12",
        "תסמן בירוק",
        "תעשה את הכותרת מודגשת",
    ],
    "clearRange": [
        "clear the contents of this range",
        "erase everything in column B",
        "clear formatting from these cells",
        "delete all data in the range",
        "נקה את התוכן של הטווח",
        "מחק הכל בעמודה B",
    ],
    "hideShow": [
        "hide column C",
        "unhide rows 5 to 10",
        "hide this sheet",
        "show hidden columns",
        "הסתר עמודה C",
        "הצג עמודות מוסתרות",
    ],
    "addComment": [
        "add a comment to cell A1",
        "insert a note in this cell",
        "annotate cell B5 with a comment",
        "הוסף הערה לתא",
        "תוסיף תגובה בתא A1",
    ],
    "addHyperlink": [
        "add a link to cell A1",
        "insert a hyperlink",
        "link this cell to a URL",
        "הוסף קישור לתא",
        "תוסיף היפרלינק",
    ],
    "groupRows": [
        "group rows 3 through 8",
        "collapse these rows",
        "create an outline group",
        "ungroup columns B to E",
        "קבץ שורות 3 עד 8",
        "כווץ את השורות האלה",
    ],
    "setRowColSize": [
        "set row height to 30",
        "make column A wider",
        "change column width to 20",
        "resize row 1",
        "שנה גובה שורה ל-30",
        "הרחב את עמודה A",
    ],
    "copyPasteRange": [
        "copy this range and paste it to Sheet2",
        "duplicate these cells to another location",
        "paste values only from A1:B10 to D1",
        "copy formatting from one range to another",
        "העתק את הטווח לגיליון 2",
        "הדבק ערכים בלבד",
    ],
    "pageLayout": [
        "set the page to landscape orientation",
        "change margins to 1 inch on all sides",
        "set print area to A1:G20",
        "hide gridlines on this sheet",
        "set paper size to A4",
        "שנה לדף לרוחב",
        "הגדר אזור הדפסה",
    ],
    "insertPicture": [
        "insert an image into the sheet",
        "add a logo picture at the top left",
        "embed a base64 image at position 100,50",
        "הוסף תמונה לגיליון",
        "שים לוגו בפינה",
    ],
    "insertShape": [
        "insert a rectangle shape",
        "add a red arrow pointing right",
        "draw an oval with blue fill",
        "insert a star shape at position 200,100",
        "הוסף צורה מלבנית",
        "שים חץ אדום",
    ],
    "insertTextBox": [
        "add a text box with the title 'Summary'",
        "insert a text box at the top of the sheet",
        "create a text box with bold 14pt text",
        "הוסף תיבת טקסט עם כותרת",
        "שים תיבת טקסט בראש הדף",
    ],
    "addSlicer": [
        "add a slicer for the pivot table",
        "create a slicer to filter by Region",
        "add a slicer for the sales table by category",
        "הוסף כלי סינון לטבלת הציר",
        "תוסיף סלייסר לפי אזור",
    ],
    "splitColumn": [
        "split the full name column into first and last name",
        "break this column apart by commas",
        "split email column on the @ sign",
        "separate city and state into two columns",
        "פצל את עמודת השם לשם פרטי ומשפחה",
        "הפרד את העמודה לפי פסיקים",
    ],
    "unpivot": [
        "unpivot the monthly columns into a tall table",
        "melt wide format into long format",
        "reshape from wide to tall",
        "turn these year columns into rows",
        "הפוך עמודות חודשיות לשורות",
        "המר מרחב לאורך",
    ],
    "crossTabulate": [
        "cross-tabulate region by product",
        "build a contingency table of category and status",
        "count occurrences of A vs B",
        "make a cross-tab matrix",
        "טבלה צולבת של אזור מול מוצר",
        "ספור הצלבות בין קטגוריות",
    ],
    "bulkFormula": [
        "apply this formula to the whole column",
        "fill this formula down for every row",
        "add this formula to all rows in the data",
        "החל את הנוסחה על כל העמודה",
        "מלא את הנוסחה למטה לכל השורות",
    ],
    "compareSheets": [
        "compare Sheet1 and Sheet2 and show differences",
        "find cells that differ between these two ranges",
        "diff the old and new versions",
        "highlight what changed between the sheets",
        "השווה בין שני גיליונות",
        "מצא הבדלים בין הגרסאות",
    ],
    "consolidateRanges": [
        "combine these three ranges into one table",
        "stack data from multiple sheets together",
        "merge these ranges vertically",
        "consolidate data from Q1, Q2, Q3, Q4 sheets",
        "אחד את הטווחים לטבלה אחת",
        "ערום נתונים ממספר גיליונות",
    ],
    "extractPattern": [
        "extract email addresses from this column",
        "pull phone numbers out of the messy text",
        "get the URLs from column B",
        "extract dates from the description column",
        "חלץ כתובות אימייל מהעמודה",
        "שלוף מספרי טלפון מהטקסט",
    ],
    "categorize": [
        "categorize rows as corporate or personal",
        "tag each row based on these rules",
        "classify customers into buckets",
        "label each amount as small/medium/large",
        "סווג שורות לפי קטגוריה",
        "תייג כל שורה לפי כללים",
    ],
    "fillBlanks": [
        "fill empty cells with the value above",
        "forward-fill the merged category column",
        "fill blanks with zero",
        "carry values down to fill empty rows",
        "מלא תאים ריקים עם הערך מלמעלה",
        "השלם ריקים באפס",
    ],
    "subtotals": [
        "add subtotals by department",
        "insert subtotal rows for each category",
        "group with subtotals by region",
        "create Excel subtotals grouped by product",
        "הוסף סיכומי ביניים לפי מחלקה",
        "שורות סיכום לכל קטגוריה",
    ],
    "transpose": [
        "transpose this range",
        "flip rows and columns",
        "swap rows with columns",
        "paste special transpose",
        "הפוך שורות לעמודות",
        "טרנספוז את הטווח",
    ],
    "namedRange": [
        "name this range SalesData",
        "create a named range",
        "give this range a name",
        "define a name for these cells",
        "תן שם לטווח הזה",
        "צור טווח בעל שם",
    ],
    "fuzzyMatch": [
        "fuzzy match company names between two columns",
        "approximate string matching",
        "find similar names even if spelled differently",
        "match IBM with I.B.M.",
        "התאמה מטושטשת של שמות חברות",
        "מצא שמות דומים גם אם כתובים אחרת",
    ],
    "deleteRowsByCondition": [
        "delete all rows where column D is blank",
        "remove rows where status equals cancelled",
        "delete empty rows",
        "remove rows containing error",
        "מחק שורות ריקות",
        "הסר שורות לפי תנאי",
    ],
    "splitByGroup": [
        "split this sheet into separate sheets by department",
        "create one sheet per category",
        "separate data into tabs by column A values",
        "split by group into different worksheets",
        "פצל לגיליונות נפרדים לפי מחלקה",
        "צור גיליון לכל קטגוריה",
    ],
    "lookupAll": [
        "find all matching records not just the first",
        "VLOOKUP but return all matches",
        "lookup all instances of each value",
        "list all orders for each customer",
        "מצא את כל ההתאמות לא רק הראשונה",
        "רשום את כל ההזמנות לכל לקוח",
    ],
    "regexReplace": [
        "regex replace across this column",
        "use regular expression to clean up text",
        "replace using a pattern with capture groups",
        "extract and reformat phone numbers using regex",
        "החלפה עם ביטוי רגולרי",
        "נקה טקסט עם regex",
    ],
    "coerceDataType": [
        "convert this column from text to numbers",
        "change stored-as-text to number",
        "convert text dates to date format",
        "fix numbers stored as text",
        "המר טקסט למספרים",
        "תקן מספרים שנשמרו כטקסט",
    ],
    "normalizeDates": [
        "standardize all dates to yyyy-mm-dd",
        "fix mixed date formats in this column",
        "convert all dates to dd/mm/yyyy",
        "normalize date formats",
        "אחד את פורמט התאריכים",
        "תקן תאריכים מעורבים",
    ],
    "deduplicateAdvanced": [
        "remove duplicates but keep the most recent row",
        "deduplicate keeping the row with most data",
        "remove duplicates keep last occurrence",
        "deduplicate by name column keeping newest by date",
        "הסר כפילויות ושמור את האחרון",
        "מחק כפולים לפי תאריך עדכני",
    ],
    "joinSheets": [
        "join these two sheets like SQL LEFT JOIN",
        "merge two tables by matching ID column",
        "combine data from two sheets by key",
        "inner join sheet1 and sheet2 on column A",
        "חבר שני גיליונות לפי מפתח",
        "מזג טבלאות לפי עמודת מזהה",
    ],
    "frequencyDistribution": [
        "count how many times each value appears",
        "frequency distribution of this column",
        "create a frequency table",
        "show value counts",
        "ספור כמה פעמים כל ערך מופיע",
        "התפלגות תדירויות",
    ],
    "runningTotal": [
        "add a running total column",
        "cumulative sum of sales",
        "running balance",
        "calculate running total",
        "הוסף עמודת סכום מצטבר",
        "סכום רץ של מכירות",
    ],
    "rankColumn": [
        "rank these values from highest to lowest",
        "add a rank column",
        "rank employees by score",
        "show ranking for each row",
        "דרג את הערכים מהגבוה לנמוך",
        "הוסף עמודת דירוג",
    ],
    "topN": [
        "show me the top 10 by revenue",
        "extract bottom 5 performers",
        "top 20 products by sales",
        "get the 10 lowest values",
        "הראה לי את 10 המובילים",
        "חלץ 5 התחתונים",
    ],
    "percentOfTotal": [
        "calculate percentage of total for each row",
        "what percent does each item contribute",
        "add a percent of total column",
        "show each row as a percentage of the sum",
        "חשב אחוז מהסך לכל שורה",
        "הוסף עמודת אחוזים מהסכום",
    ],
    "growthRate": [
        "calculate month over month growth",
        "year over year growth rate",
        "period over period change",
        "growth percentage between rows",
        "חשב קצב גדילה חודשי",
        "אחוז שינוי בין תקופות",
    ],
    "consolidateAllSheets": [
        "merge all sheets into one",
        "combine every worksheet into a single sheet",
        "consolidate all tabs",
        "stack all sheets together",
        "אחד את כל הגיליונות לאחד",
        "מזג את כל הלשוניות",
    ],
    "cloneSheetStructure": [
        "copy the sheet structure without data",
        "create a blank copy with same headers and formatting",
        "duplicate the template but empty",
        "clone sheet layout only",
        "העתק מבנה גיליון בלי נתונים",
        "שכפל תבנית ריקה",
    ],
    "addReportHeader": [
        "add a title row above the data",
        "insert a report header",
        "create a formatted title at the top",
        "add a merged header row with styling",
        "הוסף כותרת דוח מעל הנתונים",
        "תוסיף שורת כותרת מעוצבת",
    ],
    "alternatingRowFormat": [
        "zebra stripe the rows",
        "alternate row colors",
        "add banded rows",
        "stripe every other row",
        "פסים לסירוגין בשורות",
        "צבע שורות מתחלפות",
    ],
    "quickFormat": [
        "format this table nicely",
        "freeze header add filters and auto-fit",
        "make this look professional",
        "apply standard table formatting",
        "עצב את הטבלה יפה",
        "תעשה את זה מקצועי",
    ],
    "refreshPivot": [
        "refresh the pivot table",
        "update pivot data",
        "recalculate the pivot",
        "refresh all pivots on this sheet",
        "רענן את טבלת הציר",
        "עדכן נתוני הפיבוט",
    ],
    "pivotCalculatedField": [
        "add a calculated field to the pivot",
        "create a profit margin field in the pivot table",
        "add a computed column to the pivot",
        "הוסף שדה מחושב לפיבוט",
        "צור שדה רווח בטבלת ציר",
    ],
    "addDropdownControl": [
        "add a dropdown list in this cell",
        "create a dropdown selector",
        "add a filter dropdown",
        "put a dropdown with options A B C",
        "הוסף רשימה נפתחת בתא",
        "צור תפריט בחירה",
    ],
    "conditionalFormula": [
        "if column A is blank use formula X otherwise formula Y",
        "apply different formulas based on a condition",
        "conditional calculation based on column value",
        "if status is active multiply by 1.1 otherwise keep same",
        "אם עמודה A ריקה תשתמש בנוסחה X",
        "חישוב מותנה לפי ערך בעמודה",
    ],
    "spillFormula": [
        "use FILTER to show only rows where column B > 100",
        "write a UNIQUE formula",
        "create a SORT formula that spills",
        "dynamic array formula with FILTER",
        "נוסחת FILTER להציג שורות מעל 100",
        "כתוב נוסחת UNIQUE",
    ],
}

_COLLECTION = "capabilities"


def init_store() -> None:
    """
    Seed the `capabilities` collection with (description, example) pairs
    for every action. Idempotent — skips if the collection is already at
    the expected count.
    """
    global _ready, _expected_count  # noqa: PLW0603

    from ..persistence.factory import get_vector_store
    from .planner import CAPABILITY_DESCRIPTIONS

    store = get_vector_store()

    ids: list[str] = []
    documents: list[str] = []
    metadatas: list[dict[str, str]] = []

    for action, description in CAPABILITY_DESCRIPTIONS.items():
        ids.append(f"{action}_desc")
        documents.append(f"{action}: {description}")
        metadatas.append({"action": action})

        for i, example in enumerate(CAPABILITY_EXAMPLES.get(action, [])):
            ids.append(f"{action}_ex_{i}")
            documents.append(example)
            metadatas.append({"action": action})

    _expected_count = len(ids)

    if store.count(_COLLECTION) == _expected_count:
        logger.info(
            "Capability store already indexed (%d docs) — skipping.",
            _expected_count,
        )
        _ready = True
        return

    # Full re-index: nuke and repopulate.
    store.recreate(_COLLECTION)
    store.upsert(_COLLECTION, ids, documents, metadatas)
    logger.info("Indexed %d capability documents.", _expected_count)
    _ready = True


def is_ready() -> bool:
    return _ready


@functools.lru_cache(maxsize=256)
def search_capabilities(query: str, top_k: int | None = None) -> list[str]:
    """
    Return the action names of the top-K most relevant capabilities.

    Uses the multilingual embedding model so English and Hebrew queries
    match against the same example pool without translation.
    """
    if not _ready:
        from .planner import CAPABILITY_DESCRIPTIONS

        return list(CAPABILITY_DESCRIPTIONS.keys())

    from ..persistence.factory import get_vector_store

    store = get_vector_store()
    k = top_k or settings.capability_top_k

    # Over-fetch to account for dedup across multiple docs per action.
    fetch = min(k * 3, max(_expected_count, 1))
    results = store.query(_COLLECTION, query, top_k=fetch)

    seen: dict[str, None] = {}
    for hit in results:
        action = (hit.get("metadata") or {}).get("action")
        if action and action not in seen:
            seen[action] = None

    return list(seen.keys())[:k]
