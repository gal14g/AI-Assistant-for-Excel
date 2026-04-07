"""
Example Store – dynamic few-shot example retrieval via vector search.

Replaces hardcoded few-shot examples with a ChromaDB-backed store that:
1. Seeds curated examples on first startup
2. Retrieves the most relevant examples for each user query
3. Promotes real user-approved interactions into the example pool

User messages are embedded into a ChromaDB collection. Full example data
(user message + assistant JSON response) lives in SQLite. At query time,
ChromaDB finds the closest user messages, then SQLite returns the full pairs.
"""

from __future__ import annotations

import logging

from ..config import settings

logger = logging.getLogger(__name__)

_collection = None
_ready = False

# ── Seed examples ────────────────────────────────────────────────────────────
# Extracted from the original hardcoded _few_shot_examples().
# Each entry: (user_message, assistant_json_response)

SEED_EXAMPLES: list[tuple[str, str]] = [
    # 1. Greeting → message
    (
        "hi",
        '{"responseType":"message","message":"Hi! I\'m your Excel assistant. Tell me what you\'d like to do with your spreadsheet.","plans":null}',
    ),
    # 2. Question → message
    (
        "what can you do?",
        '{"responseType":"message","message":"I can write formulas, create charts and pivot tables, sort and filter data, apply conditional formatting, clean text, remove duplicates, match records between sheets, manage sheets, and more. Select a range in Excel, paste it here with Ctrl+V, and tell me what you need!","plans":null}',
    ),
    # 3. Match between sheets with range tokens
    (
        "write to [[Sheet2!C:C]] the values from [[Sheet1!B:B]] where [[Sheet1!A:A]] matches [[Sheet2!A:A]]",
        '{"responseType":"plan","message":"I\'ll use XLOOKUP to match Sheet2 column A against Sheet1 column A, and write the corresponding Sheet1 column B values into Sheet2 column C.","plan":{"planId":"ex-match-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"write to Sheet2!C:C values from Sheet1!B:B where Sheet1!A:A matches Sheet2!A:A","summary":"XLOOKUP Sheet1!B into Sheet2!C based on column A match","steps":[{"id":"step_1","description":"Match Sheet2!A against Sheet1!A and write Sheet1!B values into Sheet2!C","action":"matchRecords","params":{"lookupRange":"Sheet2!A:A","sourceRange":"Sheet1!A:A","returnColumns":[2],"matchType":"exact","outputRange":"Sheet2!C:C","preferFormula":true},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 4. Chart with range token
    (
        "create a bar chart from [[Sheet1!A1:B20]]",
        '{"responseType":"plan","message":"I\'ll create a bar chart using the data in Sheet1!A1:B20.","plan":{"planId":"ex-chart-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"create a bar chart from [[Sheet1!A1:B20]]","summary":"Create bar chart from Sheet1!A1:B20","steps":[{"id":"step_1","description":"Create a bar chart from the selected range","action":"createChart","params":{"dataRange":"Sheet1!A1:B20","chartType":"bar"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 5. Pivot table — auto-detect
    (
        "create a pivot table from [[Sheet1!A1:D50]]",
        '{"responseType":"plan","message":"I\'ll create a PivotTable from the selected range. Fields will be auto-detected from the column headers.","plan":{"planId":"ex-pivot-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"create a pivot table from [[Sheet1!A1:D50]]","summary":"Create PivotTable from Sheet1!A1:D50","steps":[{"id":"step_1","description":"Create PivotTable (rows and values auto-detected from headers)","action":"createPivot","params":{"sourceRange":"Sheet1!A1:D50"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.92,"warnings":[]}}',
    ),
    # 6. Pivot table — explicit rows and values
    (
        "create a pivot from [[Sheet2!A1:D7]], set rows to [[Sheet2!A:A]] and values to sum of [[Sheet2!D:D]]",
        '{"responseType":"plan","message":"I\'ll create a PivotTable grouping by column A (rows) and summing column D values.","plan":{"planId":"ex-pivot-2","createdAt":"2024-01-01T00:00:00Z","userRequest":"create pivot from Sheet2!A1:D7, rows=Sheet2!A:A, values=sum Sheet2!D:D","summary":"PivotTable rows=col A, values=SUM col D","steps":[{"id":"step_1","description":"Create PivotTable with rows from column A and sum of column D as values","action":"createPivot","params":{"sourceRange":"Sheet2!A1:D7","rows":["Sheet2!A:A"],"values":[{"field":"Sheet2!D:D","summarizeBy":"sum"}]},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.97,"warnings":[]}}',
    ),
    # 7. Multi-step: sort → chart
    (
        "sort by column B descending then create a chart",
        '{"responseType":"plan","message":"I\'ll sort the data by column B descending, then create a chart from the sorted result.","plan":{"planId":"ex-sort-chart","createdAt":"2024-01-01T00:00:00Z","userRequest":"sort by column B descending then create a chart","summary":"Sort by column B desc then create chart","steps":[{"id":"step_1","description":"Sort data by column B descending","action":"sortRange","params":{"range":"Sheet1!A:B","sortFields":[{"columnIndex":1,"ascending":false}],"hasHeaders":true},"dependsOn":[]},{"id":"step_2","description":"Create a chart from the sorted data","action":"createChart","params":{"dataRange":"Sheet1!A:B","chartType":"columnClustered"},"dependsOn":["step_1"]}],"preserveFormatting":true,"confidence":0.9,"warnings":[]}}',
    ),
    # 8. Conditional formatting
    (
        "highlight cells in column B above 100 in red",
        '{"responseType":"plan","message":"I\'ll add a conditional formatting rule to highlight cells in column B with values greater than 100 in red.","plan":{"planId":"ex-cf-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"highlight cells in column B above 100 in red","summary":"Red highlight for Sheet1!B values > 100","steps":[{"id":"step_1","description":"Apply red fill to cells in column B where value > 100","action":"addConditionalFormat","params":{"range":"Sheet1!B:B","ruleType":"cellValue","operator":"greaterThan","values":[100],"format":{"fillColor":"#ffcccc","fontColor":"#c50f1f"}},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.93,"warnings":[]}}',
    ),
    # 9. Find & replace with Hebrew sheet name
    (
        "find and replace '19/04/2026' with '12/04/2026' in sheet חשמל",
        '{"responseType":"plan","message":"I\'ll find all occurrences of \'19/04/2026\' and replace them with \'12/04/2026\' in the חשמל sheet.","plan":{"planId":"ex-fr-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"find and replace 19/04/2026 with 12/04/2026 in sheet חשמל","summary":"Replace 19/04/2026 → 12/04/2026 in חשמל","steps":[{"id":"step_1","description":"Find and replace \'19/04/2026\' with \'12/04/2026\' in sheet חשמל","action":"findReplace","params":{"find":"19/04/2026","replace":"12/04/2026","sheetName":"חשמל"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 10. Composite multi-column match → write constant
    (
        "match [[Sheet1!C:D]] with [[Sheet2!A:B]], write 'pass' to [[Sheet1!I:I]] for matches\nReferenced ranges: [[Sheet1!C:D]], [[Sheet2!A:B]], [[Sheet1!I:I]]",
        '{"responseType":"plan","message":"I\'ll do a composite key match: comparing columns C and D on Sheet1 against columns A and B on Sheet2. For every row in Sheet1 where the combination (C, D) is found in Sheet2 (A, B), I\'ll write \'pass\' to column I.","plan":{"planId":"ex-composite-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"match Sheet1!C:D with Sheet2!A:B, write pass to Sheet1!I:I","summary":"Composite key match (C:D vs A:B), write pass to I:I","steps":[{"id":"step_1","description":"Composite key match: (C,D) in Sheet1 vs (A,B) in Sheet2 — write \'pass\' to I:I for matched rows","action":"matchRecords","params":{"lookupRange":"Sheet1!C:D","sourceRange":"Sheet2!A:B","outputRange":"Sheet1!I:I","matchType":"exact","writeValue":"pass"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.97,"warnings":[]}}',
    ),
    # 11. Hebrew sheet names — use exact name
    (
        "combine [[תוכנה!A:B]] and [[Sheet2!A:B]], write 'pass' to [[תוכנה!G:G]] wherever they match\nReferenced ranges: [[תוכנה!A:B]], [[Sheet2!A:B]], [[תוכנה!G:G]]",
        '{"responseType":"plan","message":"I\'ll do a composite key match: comparing columns A and B on sheet תוכנה against columns A and B on Sheet2. For every matched row I\'ll write \'pass\' to column G of sheet תוכנה.","plan":{"planId":"ex-hebrew-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"combine תוכנה!A:B and Sheet2!A:B, write pass to תוכנה!G:G for matches","summary":"Composite key match (A:B vs A:B), write pass to תוכנה!G:G","steps":[{"id":"step_1","description":"Composite key match: (A,B) in תוכנה vs (A,B) in Sheet2 — write \'pass\' to G:G for matched rows","action":"matchRecords","params":{"lookupRange":"תוכנה!A:B","sourceRange":"Sheet2!A:B","outputRange":"תוכנה!G:G","matchType":"exact","writeValue":"pass"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.97,"warnings":[]}}',
    ),
    # 12. Format cells
    (
        "make the header row bold and center-aligned with a blue background",
        '{"responseType":"plan","message":"I\'ll format row 1 with bold text, center alignment, and a blue background.","plan":{"planId":"ex-fmt-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"make the header row bold and center-aligned with a blue background","summary":"Format header row: bold, centered, blue fill","steps":[{"id":"step_1","description":"Apply bold, center alignment, and blue fill to the header row","action":"formatCells","params":{"range":"Sheet1!1:1","bold":true,"horizontalAlignment":"center","fillColor":"#4472C4","fontColor":"#FFFFFF"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 13. Clear range
    (
        "clear all contents in column D",
        '{"responseType":"plan","message":"I\'ll clear all contents from column D, keeping the formatting intact.","plan":{"planId":"ex-clear-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"clear all contents in column D","summary":"Clear contents of column D","steps":[{"id":"step_1","description":"Clear contents of column D","action":"clearRange","params":{"range":"Sheet1!D:D","clearType":"contents"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 14. Write formula
    (
        "add a SUM formula in B10 that totals B1:B9",
        '{"responseType":"plan","message":"I\'ll write a SUM formula in cell B10.","plan":{"planId":"ex-formula-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"add a SUM formula in B10 that totals B1:B9","summary":"SUM formula in B10","steps":[{"id":"step_1","description":"Write SUM formula in B10","action":"writeFormula","params":{"cell":"Sheet1!B10","formula":"=SUM(B1:B9)"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.98,"warnings":[]}}',
    ),
    # 15. Hide columns
    (
        "hide columns C through E",
        '{"responseType":"plan","message":"I\'ll hide columns C through E.","plan":{"planId":"ex-hide-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"hide columns C through E","summary":"Hide columns C:E","steps":[{"id":"step_1","description":"Hide columns C through E","action":"hideShow","params":{"target":"columns","rangeOrName":"C:E","hide":true},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 16. Page layout
    (
        "set landscape orientation and hide gridlines on sheet Report",
        '{"responseType":"plan","message":"Setting landscape orientation and hiding gridlines.","plan":{"summary":"Set landscape and hide gridlines","steps":[{"stepId":"s1","action":"pageLayout","params":{"sheetName":"Report","orientation":"landscape","showGridlines":false}}]}}',
    ),
    # 17. Add slicer
    (
        "add a slicer to filter the pivot table Sales by Region",
        '{"responseType":"plan","message":"Adding a slicer for the Sales pivot table filtered by Region.","plan":{"summary":"Add Region slicer","steps":[{"stepId":"s1","action":"addSlicer","params":{"sourceType":"pivotTable","sourceName":"Sales","sourceField":"Region"}}]}}',
    ),
    # 18. Create chart
    (
        "create a bar chart from range A1:C10",
        '{"responseType":"plan","message":"Creating a bar chart from the selected data.","plan":{"summary":"Create bar chart","steps":[{"id":"s1","description":"Create bar chart from A1:C10","action":"createChart","params":{"dataRange":"A1:C10","chartType":"bar"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 19. Insert/delete rows
    (
        "insert 3 empty rows above row 5",
        '{"responseType":"plan","message":"Inserting 3 rows above row 5.","plan":{"summary":"Insert rows","steps":[{"id":"s1","description":"Insert 3 rows above row 5","action":"insertDeleteRows","params":{"range":"5:7","shiftDirection":"down"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 20. Auto-fit columns
    (
        "auto-fit all column widths in the sheet",
        '{"responseType":"plan","message":"Auto-fitting column widths.","plan":{"summary":"Auto-fit columns","steps":[{"id":"s1","description":"Auto-fit all columns","action":"autoFitColumns","params":{},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 21. Add sparkline
    (
        "add sparklines in column F from data in B2:E10",
        '{"responseType":"plan","message":"Adding sparklines.","plan":{"summary":"Add line sparklines","steps":[{"id":"s1","description":"Add sparklines","action":"addSparkline","params":{"dataRange":"B2:E10","locationRange":"F2:F10","sparklineType":"line"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 22. Merge cells
    (
        "merge cells A1 to D1 into one cell",
        '{"responseType":"plan","message":"Merging cells A1:D1.","plan":{"summary":"Merge cells","steps":[{"id":"s1","description":"Merge A1:D1","action":"mergeCells","params":{"range":"A1:D1"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 23. Set number format
    (
        "format column B as currency with 2 decimal places",
        '{"responseType":"plan","message":"Formatting column B as currency.","plan":{"summary":"Set currency format","steps":[{"id":"s1","description":"Format B as currency","action":"setNumberFormat","params":{"range":"B:B","format":"$#,##0.00"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 24. Sheet operations
    (
        "add a new sheet called Summary",
        '{"responseType":"plan","message":"Adding a new sheet called Summary.","plan":{"summary":"Add sheet","steps":[{"id":"s1","description":"Add Summary sheet","action":"addSheet","params":{"sheetName":"Summary"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 25. Remove duplicates
    (
        "remove duplicate rows in range A1:F100",
        '{"responseType":"plan","message":"Removing duplicate rows.","plan":{"summary":"Remove duplicates","steps":[{"id":"s1","description":"Remove duplicates in A1:F100","action":"removeDuplicates","params":{"range":"A1:F100"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 26. Copy paste range
    (
        "copy range A1:C10 and paste it to E1",
        '{"responseType":"plan","message":"Copying range to destination.","plan":{"summary":"Copy and paste range","steps":[{"id":"s1","description":"Copy A1:C10 to E1","action":"copyPasteRange","params":{"sourceRange":"A1:C10","destinationRange":"E1","pasteType":"all"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 27. Generate data — writeValues with 2D array
    (
        "Make me a list of dates for the next week with a column to mark hours",
        '{"responseType":"plans","message":"Here are two options for creating your date list with an hours column.","plans":[{"optionLabel":"Option A: Simple date list","plan":{"planId":"gen-dates-a","createdAt":"2026-04-01T00:00:00Z","userRequest":"Make me a list of dates for the next week with a column to mark hours","summary":"Write dates for the next 7 days in column A with an empty hours column in B","steps":[{"id":"step_1","description":"Write date headers and 7 days of dates with empty hours column","action":"writeValues","params":{"range":"A1:B8","values":[["Date","Hours"],["01/04/2026",""],["02/04/2026",""],["03/04/2026",""],["04/04/2026",""],["05/04/2026",""],["06/04/2026",""],["07/04/2026",""]]},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}]}',
    ),
]


async def init_example_store() -> None:
    """
    Initialise the few-shot example ChromaDB collection and seed it.

    Safe to call multiple times — idempotent.
    """
    global _collection, _ready  # noqa: PLW0603

    from .chroma_client import get_chroma_client, get_embedding_fn
    from ..db import insert_few_shot_example

    client = get_chroma_client()
    embedding_fn = get_embedding_fn()

    _collection = client.get_or_create_collection(
        name="few_shot_examples",
        embedding_function=embedding_fn,
    )

    # Seed: insert examples into both SQLite and ChromaDB (idempotent)
    existing_ids = set(_collection.get()["ids"]) if _collection.count() > 0 else set()
    new_ids: list[str] = []
    new_docs: list[str] = []
    new_metas: list[dict[str, str]] = []

    for i, (user_msg, assistant_resp) in enumerate(SEED_EXAMPLES):
        example_id = f"seed_{i}"

        # SQLite (INSERT OR IGNORE — idempotent)
        await insert_few_shot_example(
            example_id=example_id,
            user_message=user_msg,
            assistant_response=assistant_resp,
            source="seed",
        )

        # ChromaDB — only add if not already present
        if example_id not in existing_ids:
            new_ids.append(example_id)
            new_docs.append(user_msg)
            new_metas.append({"sqlite_id": example_id, "source": "seed"})

    if new_ids:
        _collection.add(ids=new_ids, documents=new_docs, metadatas=new_metas)
        logger.info("Seeded %d new few-shot examples into ChromaDB.", len(new_ids))
    else:
        logger.info("Few-shot example store already seeded (%d examples).", _collection.count())

    _ready = True


def is_ready() -> bool:
    return _ready


async def add_user_example(
    *,
    interaction_id: str,
    user_message: str,
    assistant_response: str,
) -> None:
    """Promote an applied interaction into the few-shot example pool."""
    from ..db import insert_few_shot_example

    example_id = f"user_{interaction_id}"

    await insert_few_shot_example(
        example_id=example_id,
        user_message=user_message,
        assistant_response=assistant_response,
        source="user",
        interaction_id=interaction_id,
    )

    if _collection is not None:
        # Check if already in ChromaDB
        existing = _collection.get(ids=[example_id])
        if not existing["ids"]:
            _collection.add(
                ids=[example_id],
                documents=[user_message],
                metadatas=[{"sqlite_id": example_id, "source": "user"}],
            )
            logger.info("Promoted interaction %s as few-shot example.", interaction_id)


async def search_examples(query: str, top_k: int | None = None) -> list[dict]:
    """
    Retrieve the most relevant few-shot examples for a query.

    Returns list of {"user_message": ..., "assistant_response": ...}
    ordered by relevance.
    """
    from ..db import get_few_shot_examples_by_ids

    k = top_k or settings.few_shot_top_k

    if not _ready or _collection is None or _collection.count() == 0:
        # Fallback: return first k seed examples in order
        return [
            {"user_message": u, "assistant_response": a}
            for u, a in SEED_EXAMPLES[:k]
        ]

    results = _collection.query(
        query_texts=[query],
        n_results=min(k, _collection.count()),
    )

    if not results or not results["ids"] or not results["ids"][0]:
        return []

    matched_ids = results["ids"][0]
    return await get_few_shot_examples_by_ids(matched_ids)
