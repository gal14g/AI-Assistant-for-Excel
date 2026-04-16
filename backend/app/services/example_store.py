"""
Example Store — dynamic few-shot example retrieval via vector search.

Replaces hardcoded few-shot examples with a two-tier store:
  - `few_shot_examples` collection in the `VectorStore` holds the user
    message embeddings for nearest-neighbour lookup.
  - The `FewShotRepository` (SQLite or Postgres) stores the full pairs
    (user message + assistant JSON response) keyed by the same IDs.

At query time the vector store returns the best-matching IDs; the
repository returns the full pairs. Both the vector store and the repo
are chosen at startup via the persistence factory (Item 5).
"""

from __future__ import annotations

import logging

from ..config import settings

logger = logging.getLogger(__name__)

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
    # ── Hebrew examples ──────────────────────────────────────────────────────
    # These anchor the canonical-English + `*Localized` pattern for the LLM.
    # `message`, `summary`, `description`, `optionLabel` are ALWAYS English;
    # `messageLocalized`, `summaryLocalized`, `descriptionLocalized`,
    # `optionLabelLocalized` carry the Hebrew translation for display.
    # See the LANGUAGE RULE in chat_service.py.
    # 28. Hebrew — greeting (pure message path)
    (
        "היי מה קורה",
        '{"responseType":"message","message":"Hi! I\'m your Excel assistant. Tell me what you\'d like to do with your spreadsheet.","messageLocalized":"היי! אני העוזר שלך ל-Excel. ספר לי מה תרצה לעשות בגיליון האלקטרוני שלך.","plans":null}',
    ),
    # 29. Hebrew — sort descending
    (
        "מיין את [[גיליון1!A:B]] לפי עמודה B בסדר יורד",
        '{"responseType":"plan","message":"I\'ll sort the data in גיליון1!A:B by column B in descending order.","messageLocalized":"אמיין את הנתונים בגיליון1!A:B לפי עמודה B בסדר יורד.","plan":{"planId":"ex-he-sort","createdAt":"2026-04-01T00:00:00Z","userRequest":"מיין את גיליון1!A:B לפי עמודה B בסדר יורד","summary":"Sort גיליון1!A:B by column B descending","summaryLocalized":"מיון גיליון1!A:B לפי עמודה B בסדר יורד","steps":[{"id":"step_1","description":"Sort גיליון1!A:B by column B descending","descriptionLocalized":"מיין את גיליון1!A:B לפי עמודה B בסדר יורד","action":"sortRange","params":{"range":"גיליון1!A:B","sortFields":[{"columnIndex":1,"ascending":false}],"hasHeaders":true},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 30. Hebrew — pivot + chart chaining
    (
        "צור טבלת ציר מ-[[מכירות!A1:D500]] עם שורות לפי אזור וסכום מכירות, ואז צור תרשים",
        '{"responseType":"plan","message":"I\'ll create a PivotTable from the מכירות sheet with rows by region and SUM of sales as values, then build a column chart from it.","messageLocalized":"אצור טבלת ציר מגיליון מכירות כשהשורות הן האזור והערכים הם סכום המכירות, ולאחר מכן אצור תרשים עמודות.","plan":{"planId":"ex-he-pivot-chart","createdAt":"2026-04-01T00:00:00Z","userRequest":"טבלת ציר + תרשים ממכירות","summary":"PivotTable by region → column chart","summaryLocalized":"טבלת ציר לפי אזור → תרשים עמודות","steps":[{"id":"step_1","description":"Create PivotTable: rows=region, values=SUM(sales)","descriptionLocalized":"צור טבלת ציר: שורות=אזור, ערכים=SUM(מכירות)","action":"createPivot","params":{"sourceRange":"מכירות!A1:D500","rows":["אזור"],"values":[{"field":"מכירות","summarizeBy":"sum"}]},"dependsOn":[]},{"id":"step_2","description":"Create a column chart from the PivotTable","descriptionLocalized":"צור תרשים עמודות מטבלת הציר","action":"createChart","params":{"dataRange":"{{step_1.outputRange}}","chartType":"columnClustered"},"dependsOn":["step_1"]}],"preserveFormatting":true,"confidence":0.92,"warnings":[]}}',
    ),
    # 31. Hebrew — find-replace date format
    (
        "החלף את כל המופעים של '31/12/2025' ב-'01/01/2026' בגיליון חשבוניות",
        '{"responseType":"plan","message":"I\'ll replace every occurrence of \'31/12/2025\' with \'01/01/2026\' in the חשבוניות sheet.","messageLocalized":"אחליף את כל המופעים של \'31/12/2025\' ב-\'01/01/2026\' בגיליון חשבוניות.","plan":{"planId":"ex-he-fr","createdAt":"2026-04-01T00:00:00Z","userRequest":"החלף 31/12/2025 → 01/01/2026 בחשבוניות","summary":"Replace 31/12/2025 → 01/01/2026 in חשבוניות","summaryLocalized":"החלפת 31/12/2025 → 01/01/2026 בגיליון חשבוניות","steps":[{"id":"step_1","description":"Find and replace \'31/12/2025\' with \'01/01/2026\' in חשבוניות","descriptionLocalized":"מצא והחלף \'31/12/2025\' ב-\'01/01/2026\' בחשבוניות","action":"findReplace","params":{"find":"31/12/2025","replace":"01/01/2026","sheetName":"חשבוניות"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 32. Hebrew — conditional format
    (
        "צבע באדום את כל התאים בעמודה [[דוח!C:C]] שגדולים מ-1000",
        '{"responseType":"plan","message":"I\'ll add a conditional-formatting rule that colors cells red in column C of the דוח sheet when the value is greater than 1000.","messageLocalized":"אוסיף כלל עיצוב מותנה שצובע באדום את התאים בעמודה C בגיליון דוח כאשר הערך גדול מ-1000.","plan":{"planId":"ex-he-cf","createdAt":"2026-04-01T00:00:00Z","userRequest":"צבע באדום תאים בדוח!C:C מעל 1000","summary":"Red conditional format for דוח!C:C > 1000","summaryLocalized":"עיצוב מותנה אדום לתאים > 1000 בדוח!C:C","steps":[{"id":"step_1","description":"Apply red fill when C > 1000","descriptionLocalized":"החל עיצוב מותנה אדום כאשר C > 1000","action":"addConditionalFormat","params":{"range":"דוח!C:C","ruleType":"cellValue","operator":"greaterThan","values":[1000],"format":{"fillColor":"#ffcccc","fontColor":"#c50f1f"}},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.93,"warnings":[]}}',
    ),
    # 33. Hebrew — formula filldown
    (
        "הוסף נוסחת SUM בעמודה D שמחברת את A, B, ו-C ומלא עד שורה 100",
        '{"responseType":"plan","message":"I\'ll write =SUM(A2:C2) into D2 and fill it down to row 100.","messageLocalized":"אכתוב נוסחת SUM ב-D2 שמסכמת A2:C2, ואמלא אותה עד שורה 100.","plan":{"planId":"ex-he-sum","createdAt":"2026-04-01T00:00:00Z","userRequest":"SUM(A:C) ב-D2 עד שורה 100","summary":"SUM formula in column D, filled to row 100","summaryLocalized":"נוסחת SUM בעמודה D עם מילוי עד שורה 100","steps":[{"id":"step_1","description":"Write =SUM(A2:C2) in D2 and fill down to row 100","descriptionLocalized":"כתוב =SUM(A2:C2) ב-D2 ומלא עד שורה 100","action":"writeFormula","params":{"cell":"D2","formula":"=SUM(A2:C2)","fillDown":98},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.97,"warnings":[]}}',
    ),
    # 34. Hebrew — text cleanup chain
    (
        "נקה רווחים מיותרים ב-[[לקוחות!B:B]] ואז הסר כפילויות לפי עמודה B",
        '{"responseType":"plan","message":"I\'ll clean whitespace in column B, then remove duplicate rows by that same column.","messageLocalized":"אנקה רווחים בעמודה B ולאחר מכן אסיר שורות כפולות לפי אותה עמודה.","plan":{"planId":"ex-he-clean","createdAt":"2026-04-01T00:00:00Z","userRequest":"ניקוי רווחים + הסרת כפילויות בלקוחות!B","summary":"Clean column B, then remove duplicates","summaryLocalized":"ניקוי עמודה B ואז הסרת כפילויות","steps":[{"id":"step_1","description":"Clean whitespace in column B","descriptionLocalized":"נקה רווחים מיותרים בעמודה B","action":"cleanupText","params":{"range":"לקוחות!B:B","operations":["trim","normalizeWhitespace"]},"dependsOn":[]},{"id":"step_2","description":"Remove duplicate rows by column B","descriptionLocalized":"הסר שורות כפולות לפי עמודה B","action":"removeDuplicates","params":{"range":"לקוחות!A:Z","columnIndexes":[1]},"dependsOn":["step_1"]}],"preserveFormatting":true,"confidence":0.93,"warnings":[]}}',
    ),
    # 35. Hebrew — duplicate-sidecar layout (lateralSpreadDuplicates)
    (
        "כל ערך שמופיע ב-[[Sheet5!B:B]] יותר מפעם אחת, תעביר את כל השורה שלו להיות משמאל לשורה הראשונה בה מופיע",
        '{"responseType":"plan","message":"I\'ll lay every duplicate of column B next to its first occurrence horizontally — the original first-occurrence rows stay where they are, and each subsequent duplicate gets lifted out of its vertical position and pasted as a new block of columns on the left of the anchor row. Single-step, no temp sheets.","messageLocalized":"אפרוש כל שורה שערך עמודה B שלה חוזר — השורה הראשונה של כל ערך נשארת במקומה, וכל שורה כפולה עוברת להיות בבלוק של עמודות משמאל לשורה הראשונה שלה. צעד אחד, בלי גיליון עזר.","plan":{"planId":"ex-he-sidecar","createdAt":"2026-04-01T00:00:00Z","userRequest":"duplicate sidecar on Sheet5!B","summary":"Lateral-spread duplicates of Sheet5!B next to first occurrence","summaryLocalized":"פרישה צידית של כפילויות עמודה B בגיליון5 ליד המופע הראשון","steps":[{"id":"step_1","description":"Spread duplicates of column B horizontally, left of each first occurrence","descriptionLocalized":"פרוש את כפילויות עמודה B אופקית, משמאל לכל מופע ראשון","action":"lateralSpreadDuplicates","params":{"sourceRange":"Sheet5!A:Z","keyColumnIndex":1,"hasHeaders":true,"direction":"left","removeOriginalDuplicates":true},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 36. within-row match extraction (extractMatchedToNewRow)
    (
        "in [[Sheet1!A1:F100]] when column name equals column name2, move number2 and price2 to a new row below with the shared name",
        '{"responseType":"plan","message":"I\'ll split each row where name equals name2 into two rows: the first keeps the primary side (name, number, price); the new row below it carries the shared name plus number2 and price2. One step, one sync.","plan":{"planId":"ex-extract-match","createdAt":"2026-04-01T00:00:00Z","userRequest":"split rows where name==name2 into two rows","summary":"Extract matched-name rows into a new row below","steps":[{"id":"step_1","description":"When column name (0) equals column name2 (3), move columns number2 (4) and price2 (5) to a new row below, duplicating the name value","action":"extractMatchedToNewRow","params":{"sourceRange":"Sheet1!A1:F100","keyColumnIndexA":0,"keyColumnIndexB":3,"extractColumnIndexes":[4,5],"hasHeaders":true,"caseSensitive":false},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.93,"warnings":[]}}',
    ),
    # 37. reorderRows — move matching to top
    (
        "move all rows in [[Sheet1!A1:D200]] where column C equals 'urgent' to the top",
        '{"responseType":"plan","message":"I\'ll reorder the rows so every row with \'urgent\' in column C sits at the top, with the rest keeping their original order.","plan":{"planId":"ex-reorder","createdAt":"2026-04-01T00:00:00Z","userRequest":"move urgent rows to top","summary":"Move rows with column C == \'urgent\' to the top","steps":[{"id":"step_1","description":"Reorder rows — move matching to top","action":"reorderRows","params":{"range":"Sheet1!A1:D200","mode":"moveMatching","conditionColumn":2,"condition":"equals","conditionValue":"urgent","destination":"top","hasHeaders":true},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}',
    ),
    # 38. fillSeries — numbering
    (
        "number rows 1 to 100 in column A of [[Sheet1!A1:A100]]",
        '{"responseType":"plan","message":"I\'ll write the sequence 1..100 into A1:A100.","plan":{"planId":"ex-fillseries","createdAt":"2026-04-01T00:00:00Z","userRequest":"number 1-100 in column A","summary":"fillSeries 1..100 in A1:A100","steps":[{"id":"step_1","description":"Fill number series 1..100","action":"fillSeries","params":{"range":"Sheet1!A1:A100","seriesType":"number","start":1,"step":1},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.98,"warnings":[]}}',
    ),
    # 39. Hebrew — RTL sheet direction
    (
        "הפוך את הגיליון הזה לימין-לשמאל",
        '{"responseType":"plan","message":"I\'ll request RTL display for the active sheet. Office.js has no API for this — the handler will return a warning and you\'ll need to toggle View > Sheet Right-to-Left manually.","messageLocalized":"אבקש תצוגת ימין-לשמאל לגיליון הפעיל. ל-Office.js אין API לכך — המערכת תחזיר אזהרה ותצטרך להפעיל ידנית דרך תצוגה > גיליון מימין לשמאל.","plan":{"planId":"ex-rtl","createdAt":"2026-04-01T00:00:00Z","userRequest":"make this sheet RTL","summary":"Request RTL direction for the active sheet","summaryLocalized":"בקשת כיוון ימין-לשמאל לגיליון הפעיל","steps":[{"id":"step_1","description":"Set sheet direction to RTL","descriptionLocalized":"הגדר את כיוון הגיליון לימין-לשמאל","action":"setSheetDirection","params":{"direction":"rtl"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.9,"warnings":[]}}',
    ),
    # 40. Hebrew — self-join row-relation comparison (LET + array-MATCH).
    # Teaches the pattern: "for each row, find a DIFFERENT row of the same
    # group whose column value has a relationship to this row's value (e.g.
    # previous-day end date), then compare a third column between the two".
    # This is the pattern the LLM consistently fumbled into XLOOKUP / single-
    # key VLOOKUP / or in-row comparison; the correct shape is
    # IFERROR(LET(MATCH(1, array-criteria, 0), INDEX(...))).
    (
        "לכל כפילות שיש כאן: [[Sheet6!A:A]] , אם הערך כאן: [[Sheet6!B:B]] = לערך כאן [[Sheet6!C:C]] פחות אחד, וגם בשורה של הערך פחות אחד בעמודה הזו: [[Sheet6!D:D]] מופיע ״נציבותי״ ואז בשורה של הערך מופיע משהו שלא שווה ״נציבותי״ אז תרשום בעמודה e עזיבה, אחרת אם מופיע משהו שלא שווה ״נציבותי״ ואז בשורה של הערך מופיע ״נציבותי״ תרשום בעמודה e קליטה",
        '{"responseType":"plan","message":"Self-join comparison: for every row I find a PREVIOUS row of the same key (column A) whose end-date (C) is one day before this row\'s start-date (B), then compare column D between the two rows. The formula uses IFERROR(LET(MATCH(1, multi-criteria, 0), INDEX(...))) — array criteria let MATCH locate the previous row by two conditions at once, and IFERROR turns no-match rows into empty strings instead of #N/A.","messageLocalized":"השוואת שורה-לשורה באותה קבוצה: לכל שורה אני מחפש שורה קודמת עם אותו ערך ב-A, שתאריך הסיום שלה (C) הוא יום לפני תאריך ההתחלה (B) של השורה הנוכחית, ואז משווה את עמודה D בין שתי השורות. הנוסחה משתמשת ב-IFERROR(LET(MATCH(1, קריטריון מערכי, 0), INDEX(...))) — קריטריון המערך מאפשר ל-MATCH למצוא את השורה הקודמת לפי שני תנאים בו-זמנית, ו-IFERROR גורם לשורות ללא התאמה להחזיר ריק במקום #N/A.","plan":{"planId":"ex-he-selfjoin","createdAt":"2026-04-01T00:00:00Z","userRequest":"self-join עזיבה/קליטה בעמודה E של Sheet6","summary":"Write self-join comparison formula into Sheet6!E2 and fill down","summaryLocalized":"כתיבת נוסחת השוואה self-join ב-Sheet6!E2 עם מילוי כלפי מטה","steps":[{"id":"step_1","description":"Write the self-join formula into E2 and fill down. Uses LET+MATCH with multi-criteria array (A matches + C=B-1), then compares the previous row\'s D against the current row\'s D. IFERROR wraps the lookup so rows without a predecessor return \\"\\".","descriptionLocalized":"כתוב נוסחת self-join ב-E2 ומלא כלפי מטה. NOTE מ-LET+MATCH עם קריטריון מערכי (A זהה + C=B-1), ואז משווה את D של השורה הקודמת ל-D של השורה הנוכחית. IFERROR עוטף את החיפוש כדי שורות ללא קודמת יחזירו ריק.","action":"writeFormula","params":{"cell":"Sheet6!E2","formula":"=IFERROR(LET(prevIdx,MATCH(1,(A$2:A$100=A2)*(C$2:C$100=B2-1),0),prevD,INDEX(D$2:D$100,prevIdx),IF(AND(prevD=\\"נציבותי\\",D2<>\\"נציבותי\\"),\\"עזיבה\\",IF(AND(prevD<>\\"נציבותי\\",D2=\\"נציבותי\\"),\\"קליטה\\",\\"\\"))),\\"\\")","fillDown":98},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.9,"warnings":["Formula uses LET + INDEX/MATCH array criteria. On Excel 365 it runs natively. On Excel 2016/2019 the writeFormula handler auto-rewrites LET → inlined bindings and INDEX(range, MATCH(1, array, 0)) → LOOKUP(2, 1/array, range), producing an equivalent formula that works without Ctrl+Shift+Enter.","Adjust the A$2:A$100 / C$2:C$100 / D$2:D$100 bounds to the real data range before or after writing.","Assumes B and C are native date-typed cells (not text). If dates are text-typed, the C=B-1 arithmetic fails silently, producing #N/A."]}}',
    ),
]


_COLLECTION = "few_shot_examples"

# Bump this whenever SEED_EXAMPLES contents change in a way that should
# override previously-seeded rows (e.g. a prompt-format migration like the
# canonical-English + `*Localized` rewrite). `init_example_store` detects
# older versions present in the vector store and purges them before
# re-seeding, so the LLM never sees stale (mis-formatted) few-shots.
_SEED_VERSION = "v4"


async def init_example_store() -> None:
    """
    Seed the few-shot example pool in both the repository (full pairs) and
    the vector store (user-message embeddings). Idempotent.
    """
    global _ready  # noqa: PLW0603

    from ..persistence.factory import get_repositories, get_vector_store

    repos = get_repositories()
    store = get_vector_store()

    # ── Migration: purge any earlier-version seeds from the vector store ──
    # Pre-v2 the IDs were `seed_0`, `seed_1`, … (no version). Pre-v3 used
    # `seed_v2_…`. Any time _SEED_VERSION bumps, we want the old rows gone
    # so retrieval can't resurrect stale few-shots. The check is a bounded
    # ID scan (enough to cover every historical seed), collect matches,
    # delete. Orphaned DB rows are harmless: retrieval goes through the
    # vector store first, so DB entries not referenced from vectors are
    # never surfaced.
    unversioned_ids = [f"seed_{i}" for i in range(100)]
    prior_versions = [f"seed_v{v}_{i}" for v in ("2", "3") for i in range(100)]
    candidate_stale = unversioned_ids + prior_versions
    legacy_found = store.get_by_ids(_COLLECTION, candidate_stale)
    if legacy_found:
        stale_ids = [r["id"] for r in legacy_found]
        store.delete(_COLLECTION, stale_ids)
        logger.info("Removed %d pre-%s few-shot seed(s).", len(stale_ids), _SEED_VERSION)

    # Vector store: figure out which seed IDs are already present by
    # fetching the list and skipping duplicates. We cap seeding at the
    # length of SEED_EXAMPLES so re-seeding is cheap on hot restarts.
    seed_ids = [f"seed_{_SEED_VERSION}_{i}" for i in range(len(SEED_EXAMPLES))]
    existing = store.get_by_ids(_COLLECTION, seed_ids)
    existing_ids = {r["id"] for r in existing}

    new_ids: list[str] = []
    new_docs: list[str] = []
    new_metas: list[dict[str, str]] = []

    for i, (user_msg, assistant_resp) in enumerate(SEED_EXAMPLES):
        example_id = seed_ids[i]

        # Repo (INSERT OR IGNORE — idempotent)
        await repos.few_shot.insert(
            example_id=example_id,
            user_message=user_msg,
            assistant_response=assistant_resp,
            source="seed",
        )

        if example_id not in existing_ids:
            new_ids.append(example_id)
            new_docs.append(user_msg)
            new_metas.append({"sqlite_id": example_id, "source": "seed"})

    if new_ids:
        store.upsert(_COLLECTION, new_ids, new_docs, new_metas)
        logger.info("Seeded %d new few-shot examples.", len(new_ids))
    else:
        logger.info("Few-shot example store already seeded.")

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
    from ..persistence.factory import get_repositories, get_vector_store

    example_id = f"user_{interaction_id}"

    await get_repositories().few_shot.insert(
        example_id=example_id,
        user_message=user_message,
        assistant_response=assistant_response,
        source="user",
        interaction_id=interaction_id,
    )

    store = get_vector_store()
    existing = store.get_by_ids(_COLLECTION, [example_id])
    if not existing:
        store.upsert(
            _COLLECTION,
            [example_id],
            [user_message],
            [{"sqlite_id": example_id, "source": "user"}],
        )
        logger.info("Promoted interaction %s as few-shot example.", interaction_id)


async def search_examples(query: str, top_k: int | None = None) -> list[dict]:
    """
    Retrieve the most relevant few-shot examples for a query.

    Returns a list of {"user_message": ..., "assistant_response": ...}
    ordered by relevance.
    """
    from ..persistence.factory import get_repositories, get_vector_store

    k = top_k or settings.few_shot_top_k

    if not _ready:
        # Fallback: return first k seed examples in order.
        return [
            {"user_message": u, "assistant_response": a}
            for u, a in SEED_EXAMPLES[:k]
        ]

    store = get_vector_store()
    hits = store.query(_COLLECTION, query, top_k=k)
    if not hits:
        return []

    matched_ids = [h["id"] for h in hits]
    return await get_repositories().few_shot.get_by_ids(matched_ids)
