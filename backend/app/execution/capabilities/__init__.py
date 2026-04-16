"""
Python xlwings capability handlers — parallel to the frontend's
`engine/capabilities/` directory. One file per action, same naming.

Registration pattern mirrors the TS side:

    from app.execution.capability_registry import registry
    from app.execution.capabilities import HandlerContext

    def handler(ctx: HandlerContext, params: dict) -> dict:
        ...

    registry.register("readRange", handler, mutates=False)

All handlers take `(ctx, params)` and return a dict matching the
`StepResult` wire-format (`{status, message, outputs?, error?}`). The
executor wraps the return value, adds the step id, and handles timing.

This file force-imports every handler module so that simply doing
`import app.execution.capabilities` makes the registry fully populated.
Keep the list alphabetized.
"""

# Core read/write.
from app.execution.capabilities import (  # noqa: F401
    read_range,
    write_values,
    write_formula,
    clear_range,
    copy_paste_range,
)

# Data manipulation.
from app.execution.capabilities import (  # noqa: F401
    sort_range,
    apply_filter,
    find_replace,
    insert_delete_rows,
    merge_cells,
    freeze_panes,
    auto_fit_columns,
    hide_show,
    remove_duplicates,
    fill_blanks,
    cleanup_text,
    regex_replace,
    extract_pattern,
    categorize,
    coerce_data_type,
    normalize_dates,
    split_column,
    split_by_group,
    delete_rows_by_condition,
)

# Formatting.
from app.execution.capabilities import (  # noqa: F401
    format_cells,
    set_number_format,
    set_row_col_size,
    add_conditional_format,
    add_validation,
    add_report_header,
    alternating_row_format,
    quick_format,
    add_dropdown_control,
    group_rows,
    page_layout,
)

# Sheet / workbook structure.
from app.execution.capabilities import (  # noqa: F401
    sheet_ops,
    named_range,
    clone_sheet_structure,
)

# Tables / pivots / charts.
from app.execution.capabilities import (  # noqa: F401
    create_table,
    create_pivot,
    refresh_pivot,
    pivot_calculated_field,
    create_chart,
    add_sparkline,
    add_slicer,
    add_comment,
    add_hyperlink,
)

# Drawing / media.
from app.execution.capabilities import (  # noqa: F401
    insert_picture,
    insert_shape,
    insert_text_box,
)

# Analytics / aggregation.
from app.execution.capabilities import (  # noqa: F401
    transpose,
    top_n,
    running_total,
    rank_column,
    percent_of_total,
    growth_rate,
    frequency_distribution,
    subtotals,
    unpivot,
    cross_tabulate,
    group_sum,
)

# Joins / lookups / comparisons.
from app.execution.capabilities import (  # noqa: F401
    fuzzy_match,
    deduplicate_advanced,
    join_sheets,
    lookup_all,
    match_records,
    compare_sheets,
    consolidate_ranges,
    consolidate_all_sheets,
)

# Formulas.
from app.execution.capabilities import (  # noqa: F401
    bulk_formula,
    conditional_formula,
    spill_formula,
)

# Batch 3 — row reshape + sheet ops + series generation.
from app.execution.capabilities import (  # noqa: F401
    lateral_spread_duplicates,
    extract_matched_to_new_row,
    reorder_rows,
    fill_series,
    insert_delete_columns,
    set_sheet_direction,
    tab_color,
    sheet_position,
    auto_fit_rows,
    calculation_mode,
    highlight_duplicates,
    concat_rows,
    insert_blank_rows,
)

# Batch 4 — analytical primitives.
from app.execution.capabilities import (  # noqa: F401
    tiered_formula,
    histogram,
    forecast,
    aging,
    pareto,
)
