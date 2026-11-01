"""Workbook snapshots and row/column-aligned comparison, with no GUI dependencies."""

from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime, time
from enum import Enum
from io import BytesIO
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook

from .align import align as align_sequences


class WorkbookError(Exception):
    """An actionable workbook loading error."""


class Change(str, Enum):
    DELETED = "削除"
    ADDED = "追加"
    MODIFIED = "変更"


@dataclass(frozen=True)
class CellValue:
    value: object
    kind: str

    @property
    def text(self) -> str:
        if isinstance(self.value, bool):
            return "TRUE" if self.value else "FALSE"
        if isinstance(self.value, (date, datetime, time)):
            return self.value.isoformat(sep=" ") if isinstance(self.value, datetime) else self.value.isoformat()
        return str(self.value)


@dataclass
class Sheet:
    name: str
    cells: dict[tuple[int, int], CellValue] = field(default_factory=dict)
    rows: int = 0
    columns: int = 0


@dataclass
class Workbook:
    path: Path
    sheets: dict[str, Sheet]


Pair = tuple[int | None, int | None]


@dataclass
class SheetComparison:
    """One sheet compared on an aligned display grid.

    ``row_map`` / ``col_map`` translate 1-based display rows/columns into the
    1-based Excel rows/columns of each side; ``None`` means the row/column does
    not exist on that side (it was inserted or deleted). ``changes`` holds the
    cell-level differences of paired rows and columns, keyed by display
    coordinates.
    """

    name: str
    left: Sheet | None
    right: Sheet | None
    changes: dict[tuple[int, int], Change]
    row_map: list[Pair] = field(default_factory=list)
    col_map: list[Pair] = field(default_factory=list)

    @property
    def rows(self) -> int:
        return len(self.row_map)

    @property
    def columns(self) -> int:
        return len(self.col_map)

    def source(self, side: int, row: int, column: int) -> tuple[int, int] | None:
        """Excel address on ``side`` (0 = left, 1 = right) of a display cell."""
        actual_row, actual_column = self.row_map[row - 1][side], self.col_map[column - 1][side]
        return None if actual_row is None or actual_column is None else (actual_row, actual_column)

    def row_status(self, row: int) -> Change | None:
        return _status(self.row_map[row - 1])

    def column_status(self, column: int) -> Change | None:
        return _status(self.col_map[column - 1])

    @property
    def counts(self) -> dict[Change, int]:
        """Cell-level differences of paired rows and columns."""
        return {kind: sum(value == kind for value in self.changes.values()) for kind in Change}

    @property
    def structure(self) -> dict[str, int]:
        """Inserted / deleted rows and columns."""
        return {
            f"{axis}_{kind.name.lower()}": sum(_status(pair) == kind for pair in mapping)
            for axis, mapping in (("rows", self.row_map), ("columns", self.col_map))
            for kind in (Change.ADDED, Change.DELETED)
        }

    @property
    def total(self) -> int:
        return len(self.changes) + sum(self.structure.values())

    @property
    def stops(self) -> list[tuple[int, int]]:
        """Display cells visited by 前の差分 / 次の差分: every changed cell and
        the first cell with a value of each inserted/deleted block."""
        stops = set(self.changes)
        firsts: dict[tuple[int, int], dict[int, int]] = {}
        for axis, mapping in ((0, self.row_map), (1, self.col_map)):
            previous = None
            for index, pair in enumerate(mapping, 1):
                status = _status(pair)
                if status and status != previous:
                    side = 0 if status == Change.DELETED else 1
                    if (side, axis) not in firsts:
                        firsts[side, axis] = self._first_values(side, axis)
                    other = firsts[side, axis].get(pair[side], 1)
                    stops.add((index, other) if axis == 0 else (other, index))
                previous = status
        return sorted(stops)

    def _first_values(self, side: int, axis: int) -> dict[int, int]:
        """For each Excel row (axis 0) / column (axis 1) of one side, the
        smallest display column / row holding a value."""
        sheet = (self.left, self.right)[side]
        other = self.col_map if axis == 0 else self.row_map
        display = {pair[side]: number for number, pair in enumerate(other, 1) if pair[side] is not None}
        firsts: dict[int, int] = {}
        for address in sheet.cells if sheet else ():
            number = display.get(address[1 - axis])
            if number is not None and number < firsts.get(address[axis], number + 1):
                firsts[address[axis]] = number
        return firsts


def _status(pair: Pair) -> Change | None:
    if pair[0] is None:
        return Change.ADDED
    if pair[1] is None:
        return Change.DELETED
    return None


@dataclass
class Comparison:
    left: Workbook
    right: Workbook
    sheets: list[SheetComparison]
    aligned: bool = True

    @property
    def counts(self) -> dict[Change, int]:
        totals = dict.fromkeys(Change, 0)
        for sheet in self.sheets:
            for kind, count in sheet.counts.items():
                totals[kind] += count
        return totals

    @property
    def structure(self) -> dict[str, int]:
        totals: dict[str, int] = {}
        for sheet in self.sheets:
            for key, count in sheet.structure.items():
                totals[key] = totals.get(key, 0) + count
        return totals


def read_workbook(path: str | Path, progress: Callable[[str], None] | None = None) -> Workbook:
    """Read a detached snapshot. No source file handle survives read_bytes().

    Empty strings and None are blank. Formulas are compared as expressions;
    formatting and cached calculation results do not affect comparison.
    """
    source = Path(path).expanduser().resolve()
    if source.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise WorkbookError("対応形式は .xlsx / .xlsm です。.xls は .xlsx に変換してください。")
    book = None
    try:
        with BytesIO(source.read_bytes()) as snapshot:
            book = load_workbook(snapshot, read_only=True, data_only=False, keep_links=False)
            sheets = {}
            for ws in book.worksheets:
                if progress:
                    progress(f"読み込み中: {source.name} / {ws.title}")
                # Avoid expanding massive formatted-but-empty rectangles silently.
                if (ws.max_row or 1) * (ws.max_column or 1) > 5_000_000:
                    raise WorkbookError(f"「{ws.title}」の使用範囲が大きすぎます（上限500万セル）。不要な行・列の書式を削除してください。")
                sheet = Sheet(ws.title)
                scanned = 0
                for row in ws.iter_rows():
                    scanned += len(row)
                    if scanned > 5_000_000:
                        raise WorkbookError(f"「{ws.title}」が読み込み上限500万セルを超えました。")
                    for cell in row:
                        if cell.value is None or cell.value == "":
                            continue
                        value = cell.value
                        if cell.data_type == "f" and not isinstance(value, str):
                            # openpyxl represents array/data-table formulas as objects.
                            value = repr(sorted(vars(value).items()))
                        sheet.cells[(cell.row, cell.column)] = CellValue(value, cell.data_type)
                        sheet.rows = max(sheet.rows, cell.row)
                        sheet.columns = max(sheet.columns, cell.column)
                sheets[ws.title] = sheet
            book.close()
            book = None
            return Workbook(source, sheets)
    except WorkbookError:
        raise
    except PermissionError as exc:
        raise WorkbookError(f"ファイルを読み込めません: {source.name}\nアクセス権限と、他のアプリの排他設定を確認してください。") from exc
    except FileNotFoundError as exc:
        raise WorkbookError(f"ファイルが見つかりません: {source}") from exc
    except Exception as exc:
        raise WorkbookError(f"{source.name} を読み込めません。Excel形式、破損、パスワード保護を確認してください。\n{exc}") from exc
    finally:
        if book is not None:
            book.close()


def compare_workbooks(left: Workbook, right: Workbook, align: bool = True, progress: Callable[[str], None] | None = None) -> Comparison:
    """Match sheet names, then compare cells.

    With ``align`` the rows and columns of each sheet are aligned first, so
    inserted/deleted rows and columns are detected (see ``align_sheets``).
    Without it, the same cell addresses are compared.
    """
    results = []
    for name in dict.fromkeys([*left.sheets, *right.sheets]):
        if progress:
            progress(f"差分を解析中: {name}")
        results.append(compare_sheet(name, left.sheets.get(name), right.sheets.get(name), align))
    return Comparison(left, right, results, align)


MIN_ROWS, MIN_COLUMNS = 30, 12
_SIM_BUDGET = 2_000_000
_COLUMN_SAMPLE = 1000


def compare_sheet(name: str, lhs: Sheet | None, rhs: Sheet | None, align: bool = True) -> SheetComparison:
    if align and lhs is not None and rhs is not None and lhs.cells != rhs.cells:
        row_map, col_map = align_sheets(lhs, rhs)
    else:
        rows = max(lhs.rows if lhs else 0, rhs.rows if rhs else 0)
        columns = max(lhs.columns if lhs else 0, rhs.columns if rhs else 0)
        row_map = [(r, r) for r in range(1, rows + 1)]
        col_map = [(c, c) for c in range(1, columns + 1)]
    _pad(row_map, MIN_ROWS)
    _pad(col_map, MIN_COLUMNS)
    return SheetComparison(name, lhs, rhs, _cell_changes(lhs, rhs, row_map, col_map), row_map, col_map)


def align_sheets(lhs: Sheet, rhs: Sheet) -> tuple[list[Pair], list[Pair]]:
    """Return 1-based (row_map, col_map) aligning two sheets.

    1. Rows are aligned by their values regardless of column position, so an
       inserted column does not prevent rows from matching.
    2. Columns are aligned by their values on the paired rows.
    3. Rows are aligned again cell by cell on the paired columns, which is
       more precise than step 1.
    """
    ids: dict[CellValue, int] = {}  # equal values share an int: much cheaper to hash
    a_rows, b_rows = _rows(lhs, ids), _rows(rhs, ids)
    budget = _row_budget(a_rows, b_rows)
    a_counts = [Counter(row.values()) for row in a_rows]
    b_counts = [Counter(row.values()) for row in b_rows]
    row_pairs = align_sequences(
        [frozenset(c.items()) for c in a_counts], [frozenset(c.items()) for c in b_counts],
        [frozenset(c) for c in a_counts], [frozenset(c) for c in b_counts],
        lambda i, j: _counter_similarity(a_counts[i], b_counts[j]),
        blank=frozenset(), budget=budget,
    )
    paired = [(i, j) for i, j in row_pairs if i is not None and j is not None]
    a_cols, b_cols = _columns(a_rows, lhs.columns, paired, 0), _columns(b_rows, rhs.columns, paired, 1)
    if len(paired) > _COLUMN_SAMPLE:
        step = len(paired) / _COLUMN_SAMPLE
        sample = {int(k * step) for k in range(_COLUMN_SAMPLE)}
        a_sample = [{k: v for k, v in col.items() if k in sample} for col in a_cols]
        b_sample = [{k: v for k, v in col.items() if k in sample} for col in b_cols]
    else:
        a_sample, b_sample = a_cols, b_cols
    col_pairs = align_sequences(
        [frozenset(c.items()) for c in a_cols], [frozenset(c.items()) for c in b_cols],
        [frozenset(c.values()) for c in a_cols], [frozenset(c.values()) for c in b_cols],
        lambda i, j: _dict_similarity(a_sample[i], b_sample[j]),
        blank=frozenset(), budget=max(400, _SIM_BUDGET // max(1, min(len(paired), _COLUMN_SAMPLE))),
    )
    shared = [(i, j) for i, j in col_pairs if i is not None and j is not None]
    if shared:
        a_index = {i + 1: k for k, (i, _) in enumerate(shared)}
        b_index = {j + 1: k for k, (_, j) in enumerate(shared)}
        a_cells = [{a_index[c]: v for c, v in row.items() if c in a_index} for row in a_rows]
        b_cells = [{b_index[c]: v for c, v in row.items() if c in b_index} for row in b_rows]
        row_pairs = align_sequences(
            [frozenset(r.items()) for r in a_cells], [frozenset(r.items()) for r in b_cells],
            [frozenset(r.values()) for r in a_cells], [frozenset(r.values()) for r in b_cells],
            lambda i, j: _dict_similarity(a_cells[i], b_cells[j]),
            blank=frozenset(), budget=budget,
        )
    return _one_based(row_pairs), _one_based(col_pairs)


def _rows(sheet: Sheet, ids: dict[CellValue, int]) -> list[dict[int, int]]:
    """Rows as {column: value id}."""
    rows: list[dict[int, int]] = [{} for _ in range(sheet.rows)]
    for (row, column), value in sheet.cells.items():
        rows[row - 1][column] = ids.setdefault(value, len(ids))
    return rows


def _columns(rows, count: int, paired, side: int) -> list[dict[int, int]]:
    """Column contents on the paired rows, keyed by pair index."""
    columns: list[dict[int, int]] = [{} for _ in range(count)]
    for k, pair in enumerate(paired):
        for column, value in rows[pair[side]].items():
            columns[column - 1][k] = value
    return columns


def _row_budget(a_rows, b_rows) -> int:
    cells = sum(map(len, a_rows)) + sum(map(len, b_rows))
    width = max(1, cells // max(1, len(a_rows) + len(b_rows)))
    return max(2_000, _SIM_BUDGET // width)


def _counter_similarity(a: Counter, b: Counter) -> float:
    if not a or not b:
        return 0.0
    return sum((a & b).values()) / max(a.total(), b.total())


def _dict_similarity(a: dict, b: dict) -> float:
    if not a or not b:
        return 0.0
    small, large = (a, b) if len(a) <= len(b) else (b, a)
    return sum(large.get(k) == v for k, v in small.items()) / len(large)


def _one_based(pairs) -> list[Pair]:
    return [(None if i is None else i + 1, None if j is None else j + 1) for i, j in pairs]


def _pad(mapping: list[Pair], minimum: int) -> None:
    """Extend with empty, paired rows/columns so small sheets fill the view."""
    left = max((pair[0] for pair in mapping if pair[0] is not None), default=0)
    right = max((pair[1] for pair in mapping if pair[1] is not None), default=0)
    while len(mapping) < minimum:
        left, right = left + 1, right + 1
        mapping.append((left, right))


def _cell_changes(lhs, rhs, row_map, col_map) -> dict[tuple[int, int], Change]:
    """Cell differences of paired rows × paired columns, in display coordinates."""
    a, b = _rows_by_number(lhs), _rows_by_number(rhs)
    a_column = {pair[0]: display for display, pair in enumerate(col_map, 1) if pair[0] is not None}
    b_column = {pair[1]: display for display, pair in enumerate(col_map, 1) if pair[1] is not None}
    changes = {}
    for display_row, (left_row, right_row) in enumerate(row_map, 1):
        if left_row is None or right_row is None:
            continue
        x, y = a.get(left_row, {}), b.get(right_row, {})
        for column, value in x.items():
            display_column = a_column[column]
            other = col_map[display_column - 1][1]
            if other is None:
                continue
            if other not in y:
                changes[display_row, display_column] = Change.DELETED
            elif y[other] != value:
                changes[display_row, display_column] = Change.MODIFIED
        for column in y:
            display_column = b_column[column]
            other = col_map[display_column - 1][0]
            if other is not None and other not in x:
                changes[display_row, display_column] = Change.ADDED
    return dict(sorted(changes.items()))


def _rows_by_number(sheet: Sheet | None) -> dict[int, dict[int, CellValue]]:
    rows: dict[int, dict[int, CellValue]] = {}
    if sheet:
        for (row, column), value in sheet.cells.items():
            rows.setdefault(row, {})[column] = value
    return rows


def compare_files(left: str | Path, right: str | Path, progress=None, align: bool = True) -> Comparison:
    return compare_workbooks(read_workbook(left, progress), read_workbook(right, progress), align, progress)
