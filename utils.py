from typing import Literal

def get_row_range(column: list, search: Literal["tt", "codes"]) -> tuple(int):
    searchDict = {"tt": ["Date", "Code"], "codes": ["Code", ' ']}
    row_start, row_end = 0, 0
    for cell in column:
        if cell.value == searchDict[search][0]:
            row_start = cell.row + 1
        if cell.value == searchDict[search][1]:
            row_end = cell.row - 1
    return row_start, row_end

def build_merge_map(ws):
    """Returns a dict mapping (row, col) → (master_row, master_col) for every merged cell."""
    merge_map = {}
    for merge_range in ws.merged_cells.ranges:
        master = (merge_range.min_row, merge_range.min_col)
        for row in range(merge_range.min_row, merge_range.max_row + 1):
            for col in range(merge_range.min_col, merge_range.max_col + 1):
                merge_map[(row, col)] = master
    return merge_map