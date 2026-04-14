"""Shared helpers for writing Excel outputs across requirement scripts.

These utilities centralize repeated Excel tasks so each logic script can focus
on report content rather than workbook plumbing and styling details.
"""

from pathlib import Path
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from common.result_status import STATUS_FILL_COLORS, count_statuses, coverage_formula


def default_results_path(script_file, filename):
    """Build the default output path in a sibling results folder.

    Args:
        script_file: The path to the calling logic script (usually __file__).
        filename: Target output file name, such as report.xlsx.

    Returns:
        Path object pointing to <script_dir>/results/<filename>.
    """
    return Path(script_file).resolve().parent / "results" / filename


def save_workbook(workbook, output_path):
    """Save an openpyxl workbook after ensuring output directories exist.

    Args:
        workbook: openpyxl Workbook instance to persist.
        output_path: Path-like destination for the .xlsx file.

    Returns:
        Path actually used for saving.
    """
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    return path


def thin_border(color="CCCCCC"):
    """Create a thin border on all four sides.

    Args:
        color: Hex color code used for each border edge.

    Returns:
        openpyxl Border configured with the requested color.
    """
    edge = Side(style="thin", color=color)
    return Border(left=edge, right=edge, top=edge, bottom=edge)


def alignment_center(wrap_text=True):
    """Build a centered alignment preset for table cells.

    Args:
        wrap_text: Whether cell text should wrap automatically.

    Returns:
        openpyxl Alignment object with centered horizontal/vertical positioning.
    """
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap_text)


def alignment_wrap_left():
    """Build a left-aligned, top-aligned wrapped alignment preset.

    Returns:
        openpyxl Alignment object suited for multiline descriptive text cells.
    """
    return Alignment(horizontal="left", vertical="top", wrap_text=True)


def style_header_cell(
    ws,
    row,
    col,
    value,
    *,
    fill_hex="2F5496",
    border=None,
    alignment=None,
    font_name="Arial",
    font_size=11,
    font_color="FFFFFF",
):
    """Write one styled header cell.

    Args:
        ws: openpyxl Worksheet to write into.
        row: 1-based row index.
        col: 1-based column index.
        value: Header label content.
        fill_hex: Header background color.
        border: Optional border object to apply.
        alignment: Optional alignment override.
        font_name: Header font family.
        font_size: Header font size.
        font_color: Header font color.

    Returns:
        The created/styled cell object.
    """
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(name=font_name, bold=True, size=font_size, color=font_color)
    cell.fill = PatternFill("solid", start_color=fill_hex)
    cell.alignment = alignment or alignment_center(wrap_text=True)
    if border is not None:
        cell.border = border
    return cell


def write_header_row(
    ws,
    headers,
    widths,
    *,
    row=1,
    fill_hex="2F5496",
    border=None,
    height=20,
    font_name="Arial",
    font_size=11,
    font_color="FFFFFF",
):
    """Write a complete styled header row and set column widths.

    Args:
        ws: openpyxl Worksheet to update.
        headers: Iterable of column header labels.
        widths: Iterable of column widths aligned with headers.
        row: 1-based row index for the header.
        fill_hex: Header background color.
        border: Optional border object for header cells.
        height: Row height for the header row.
        font_name: Header font family.
        font_size: Header font size.
        font_color: Header font color.

    Notes:
        headers and widths are consumed with zip, so both lists should have
        matching length to avoid accidental truncation.
    """
    for col_idx, (header, width) in enumerate(zip(headers, widths), 1):
        style_header_cell(
            ws,
            row,
            col_idx,
            header,
            fill_hex=fill_hex,
            border=border,
            font_name=font_name,
            font_size=font_size,
            font_color=font_color,
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[row].height = height


def estimate_wrapped_lines(value, col_width=60):
    """Estimate how many visual lines a value may occupy in a cell.

    This is a heuristic used for row-height sizing. It does not perfectly
    match Excel rendering but gives consistent readable results.

    Args:
        value: Cell content to estimate.
        col_width: Approximate character width used for wrapping.

    Returns:
        Integer line estimate.
    """
    text = "" if value is None else str(value)
    lines = text.split("\n")
    return sum(max(1, len(line) // col_width + 1) for line in lines)


def auto_row_height(values, col_width=60, min_height=20, line_height=19.5):
    """Estimate row height from a list of cell values.

    Args:
        values: Iterable of row cell values.
        col_width: Approximate wrapping width used by the estimator.
        min_height: Lower bound for returned row height.
        line_height: Height multiplier per estimated text line.

    Returns:
        Float/int row height suitable for Worksheet.row_dimensions[row].height.
    """
    if not values:
        return min_height
    max_lines = max(estimate_wrapped_lines(v, col_width=col_width) for v in values)
    return max(min_height, max_lines * line_height)


# ═══════════════════════════════════════════════════════════════════════════
# RESULTS-SHEET & SUMMARY-SHEET HELPERS
#
# These replace the duplicated Excel row-writing loops that appeared in
# both check_github_repo_installation_instructions and
# check_github_repo_installation_example_commands.
# ═══════════════════════════════════════════════════════════════════════════

def write_results_data_rows(
    ws,
    results,
    row_data_fn,
    *,
    status_fill_map=None,
    status_col=2,
    center_cols=None,
    border=None,
    link_cols=None,
    row_height_fn=None,
):
    """Write data rows for a Results sheet using a caller-supplied row builder.

    This extracts the boilerplate that both checkers duplicate:
    iteration, status fill, font, border, alignment, optional hyperlinks.

    Args:
        ws: openpyxl Worksheet (headers already written at row 1).
        results: List of result dicts.
        row_data_fn: Callable(result, row_number_1based) → list of cell values.
            row_number_1based starts at 1 (i.e. the first paper).
        status_fill_map: Dict mapping status str → hex color.
            Defaults to STATUS_FILL_COLORS.
        status_col: 1-based column index that should receive the status fill.
        center_cols: Set of 1-based column indices to center-align.
            Other columns get left-aligned wrap.
        border: openpyxl Border.  Uses thin_border("CCCCCC") if None.
        link_cols: Set of 1-based column indices containing hyperlinks.
        row_height_fn: Optional callable(row_values) → float height.
    """
    if status_fill_map is None:
        status_fill_map = STATUS_FILL_COLORS
    if center_cols is None:
        center_cols = {1, 2, 3}
    if border is None:
        border = thin_border("CCCCCC")
    link_cols = link_cols or set()

    center = alignment_center(wrap_text=True)
    wrap = alignment_wrap_left()

    for row_idx, r in enumerate(results, 2):
        row_data = row_data_fn(r, row_idx - 1)
        fill_color = status_fill_map.get(r.get("status"), "FFFFFF")
        row_fill = PatternFill("solid", start_color=fill_color)

        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = Font(name="Arial", size=10)
            cell.border = border
            cell.alignment = center if col_idx in center_cols else wrap
            if col_idx == status_col:
                cell.fill = row_fill
            if col_idx in link_cols and value:
                cell.hyperlink = value
                cell.font = Font(
                    name="Arial", size=10, color="0563C1", underline="single",
                )

        if row_height_fn:
            ws.row_dimensions[row_idx].height = row_height_fn(row_data)


def write_summary_sheet(
    ws,
    results,
    *,
    positive_label="Have Target Feature",
    negative_label="Missing Target Feature",
    extra_rows=None,
    fill_hex="2F5496",
    border=None,
):
    """Write a standard Metric / Value summary sheet.

    Replaces the identical summary-sheet code in both checkers.

    Args:
        ws: openpyxl Worksheet to populate.
        results: List of result dicts (used for counting).
        positive_label: Display label for 'yes' count.
        negative_label: Display label for 'no' count.
        extra_rows: Optional list of (label, value) tuples appended
            after the standard rows.
        fill_hex: Header fill color.
        border: openpyxl Border.
    """
    if border is None:
        border = thin_border("CCCCCC")
    center = alignment_center(wrap_text=True)

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 15

    style_header_cell(ws, 1, 1, "Metric", fill_hex=fill_hex, border=border, alignment=center)
    style_header_cell(ws, 1, 2, "Value", fill_hex=fill_hex, border=border, alignment=center)

    counts = count_statuses(results)
    yes = counts.get("yes", 0)
    no = counts.get("no", 0)
    skipped = counts.get("skipped", 0)
    errors = counts.get("error", 0)
    total = len(results)

    rows = [
        ("Total Repos Checked", total),
        (positive_label, yes),
        (negative_label, no),
        ("Skipped (non-GitHub)", skipped),
        ("Errors", errors),
        ("Coverage (%)", coverage_formula(yes, total)),
    ]
    if extra_rows:
        rows.extend(extra_rows)

    for r_idx, (label, value) in enumerate(rows, 2):
        ws.cell(row=r_idx, column=1, value=label).font = Font(name="Arial", size=10)
        ws.cell(row=r_idx, column=2, value=value).font = Font(name="Arial", size=10)
        ws.cell(row=r_idx, column=2).alignment = center
