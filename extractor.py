import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string


def load_excel(file_path: str):
    """Load an Excel workbook from the given path."""
    return load_workbook(filename=file_path, read_only=True, data_only=True)


def get_sheet_names(workbook) -> list[str]:
    """Return list of sheet names in the workbook."""
    return workbook.sheetnames


def parse_cell_reference(cell_ref: str) -> tuple[str, str] | tuple[str, str, str, str]:
    """
    Parse a cell reference like 'A1' or range like 'A1:B5'.
    Returns (col, row) for single cell or (col1, row1, col2, row2) for range.
    """
    cell_ref = cell_ref.strip().upper()

    if ':' in cell_ref:
        start, end = cell_ref.split(':')
        start_match = re.match(r'([A-Z]+)(\d+)', start)
        end_match = re.match(r'([A-Z]+)(\d+)', end)
        if start_match and end_match:
            return (start_match.group(1), start_match.group(2),
                    end_match.group(1), end_match.group(2))
    else:
        match = re.match(r'([A-Z]+)(\d+)', cell_ref)
        if match:
            return (match.group(1), match.group(2))

    raise ValueError(f"Invalid cell reference: {cell_ref}")


def extract_cell(sheet, cell_ref: str):
    """Extract value from a single cell."""
    parsed = parse_cell_reference(cell_ref)
    if len(parsed) == 2:
        col, row = parsed
        return sheet[f"{col}{row}"].value
    else:
        raise ValueError(f"Expected single cell, got range: {cell_ref}")


def extract_range(sheet, cell_ref: str) -> list[list]:
    """Extract values from a cell range as a 2D list."""
    parsed = parse_cell_reference(cell_ref)

    if len(parsed) == 2:
        # Single cell, return as 2D list
        col, row = parsed
        return [[sheet[f"{col}{row}"].value]]

    col1, row1, col2, row2 = parsed
    col1_idx = column_index_from_string(col1)
    col2_idx = column_index_from_string(col2)
    row1_idx = int(row1)
    row2_idx = int(row2)

    # Ensure proper ordering
    if col1_idx > col2_idx:
        col1_idx, col2_idx = col2_idx, col1_idx
    if row1_idx > row2_idx:
        row1_idx, row2_idx = row2_idx, row1_idx

    result = []
    for row in range(row1_idx, row2_idx + 1):
        row_data = []
        for col in range(col1_idx, col2_idx + 1):
            cell = sheet[f"{get_column_letter(col)}{row}"]
            row_data.append(cell.value)
        result.append(row_data)

    return result


def extract_cells(sheet, cell_refs: list[str]) -> dict:
    """
    Extract values from multiple cell references.
    Returns dict mapping cell reference to value(s).
    """
    results = {}
    for ref in cell_refs:
        ref = ref.strip()
        if ':' in ref:
            results[ref] = extract_range(sheet, ref)
        else:
            results[ref] = extract_cell(sheet, ref)
    return results
