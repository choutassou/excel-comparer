from pathlib import Path
import random
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook as ExcelWorkbook, load_workbook
from source.core import CellValue, Change, Comparison, Sheet, Workbook, WorkbookError, compare_files, compare_sheet, read_workbook


def assert_consistent(test, comparison):
    """Every value of both sides is shown exactly once, and the reported
    changes are exactly the differences between paired display cells."""
    for sheet in comparison.sheets:
        for side, source in enumerate((sheet.left, sheet.right)):
            for axis, mapping in enumerate((sheet.row_map, sheet.col_map)):
                numbers = [pair[side] for pair in mapping if pair[side] is not None]
                test.assertEqual(numbers, sorted(set(numbers)))
                if source:
                    test.assertLessEqual(source.columns if axis else source.rows, max(numbers))
        expected = {}
        for row, (lr, rr) in enumerate(sheet.row_map, 1):
            for col, (lc, rc) in enumerate(sheet.col_map, 1):
                if None in (lr, rr, lc, rc):
                    continue
                x = sheet.left.cells.get((lr, lc)) if sheet.left else None
                y = sheet.right.cells.get((rr, rc)) if sheet.right else None
                if x != y:
                    expected[row, col] = Change.ADDED if x is None else Change.DELETED if y is None else Change.MODIFIED
        test.assertEqual(sheet.changes, expected)


def grid_sheet(grid):
    sheet = Sheet("s")
    for r, row in enumerate(grid, 1):
        for c, value in enumerate(row, 1):
            if value is not None:
                sheet.cells[r, c] = CellValue(value, "n" if isinstance(value, int) else "s")
                sheet.rows, sheet.columns = max(sheet.rows, r), max(sheet.columns, c)
    return sheet


TABLE = [["ID", "名前", "部署", "年齢"]] + [[i, f"name{i}", f"dept{i % 3}", 20 + i] for i in range(1, 9)]


class AlignmentTests(unittest.TestCase):
    def compare(self, left, right, align=True):
        result = compare_sheet("s", grid_sheet(left), grid_sheet(right), align)
        return result, result.row_map, result.col_map

    def test_rows_and_columns_inserted_deleted_and_edited_together(self):
        right = [list(row) for row in TABLE]
        right.insert(3, [100, "new", "deptX", 99])
        del right[7]
        for k, row in enumerate(right):
            row.insert(2, "備考" if k == 0 else f"memo{k}")
        right[5][4] = "EDITED"
        result, rows, cols = self.compare(TABLE, right)
        self.assertEqual(rows[:11], [(1, 1), (2, 2), (3, 3), (None, 4), (4, 5), (5, 6), (6, 7), (7, None), (8, 8), (9, 9), (10, 10)])
        self.assertEqual(cols[:5], [(1, 1), (2, 2), (None, 3), (3, 4), (4, 5)])
        self.assertEqual(result.changes, {(6, 5): Change.MODIFIED})
        self.assertEqual(result.structure, {"rows_added": 1, "rows_deleted": 1, "columns_added": 1, "columns_deleted": 0})
        self.assertEqual(result.source(0, 6, 5), (5, 4))
        self.assertEqual(result.source(1, 6, 5), (6, 5))
        self.assertIsNone(result.source(0, 4, 1))
        self.assertEqual((result.row_status(4), result.row_status(8), result.column_status(3)), (Change.ADDED, Change.DELETED, Change.ADDED))
        self.assertEqual(result.stops, [(1, 3), (4, 1), (6, 5), (8, 1)])

    def test_rows_inserted_at_top_and_column_deleted(self):
        right = [["title"], []] + [[row[0], row[1], row[3]] for row in TABLE]
        result, rows, cols = self.compare(TABLE, right)
        self.assertEqual(rows[:4], [(None, 1), (None, 2), (1, 3), (2, 4)])
        self.assertEqual(cols[:4], [(1, 1), (2, 2), (3, None), (4, 3)])
        self.assertEqual(result.changes, {})
        self.assertEqual(result.stops, [(1, 1), (3, 3)])

    def test_rewritten_rows_are_modifications_not_shifts(self):
        result, rows, cols = self.compare([["a", "b"], ["c", "d"]], [["w", "x"], ["y", "z"]])
        self.assertEqual(rows[:2], [(1, 1), (2, 2)])
        self.assertEqual(result.counts[Change.MODIFIED], 4)
        self.assertFalse(any(result.structure.values()))

    def test_edited_row_next_to_inserted_row(self):
        right = [list(row) for row in TABLE]
        right[4][1] = "renamed"
        right.insert(4, [50, "inserted", "dept9", 77])
        result, rows, _ = self.compare(TABLE, right)
        self.assertEqual(rows[3:7], [(4, 4), (None, 5), (5, 6), (6, 7)])
        self.assertEqual(result.changes, {(6, 2): Change.MODIFIED})

    def test_blank_rows_and_duplicates(self):
        left = [["x", 1], [None], ["x", 1], ["y", 2], [None], ["x", 1]]
        right = [["x", 1], [None], ["new", 0], ["x", 1], ["y", 2], [None], ["x", 1]]
        result, rows, _ = self.compare(left, right)
        self.assertEqual(rows[:7], [(1, 1), (2, 2), (None, 3), (3, 4), (4, 5), (5, 6), (6, 7)])
        self.assertEqual(result.changes, {})

    def test_position_mode(self):
        right = [["title"]] + TABLE
        result, rows, cols = self.compare(TABLE, right, align=False)
        self.assertTrue(all(left == right for left, right in rows + cols))
        self.assertGreater(result.counts[Change.MODIFIED], 30)
        self.assertFalse(any(result.structure.values()))

    def test_random_edits_insertions_and_deletions(self):
        rng = random.Random(20261101)
        for _ in range(60):
            width, height = rng.randint(2, 9), rng.randint(2, 30)
            # First row/column always hold values, so no row/column lies outside the used range.
            base = [[f"{r}-{c}-{rng.randint(0, 3)}" if rng.random() < .85 or 0 in (r, c) else None for c in range(width)] for r in range(height)]
            kind = rng.choice(("insert", "delete", "edit", "mixed"))
            right = [list(row) for row in base]
            inserted_rows = inserted_cols = deleted_rows = deleted_cols = edits = 0
            if kind in ("insert", "mixed"):
                for _ in range(rng.randint(1, 3)):
                    right.insert(rng.randint(0, len(right)), [f"new{rng.random()}" for _ in range(len(right[0]))])
                    inserted_rows += 1
                for _ in range(rng.randint(0, 2)):
                    position = rng.randint(0, len(right[0]))
                    for row in right:
                        row.insert(position, f"col{rng.random()}")
                    inserted_cols += 1
            if kind in ("delete", "mixed"):
                for _ in range(rng.randint(1, min(3, len(right) - 1))):
                    del right[rng.randrange(len(right))]
                    deleted_rows += 1
                if len(right[0]) > 2 and rng.random() < .5:
                    position = rng.randrange(len(right[0]))
                    for row in right:
                        del row[position]
                    deleted_cols += 1
            if kind in ("edit", "mixed"):
                for _ in range(rng.randint(1, 3)):
                    row = rng.choice([row for row in right if any(v is not None for v in row)])
                    column = rng.choice([c for c, v in enumerate(row) if v is not None])
                    row[column] = f"edited{rng.random()}"
                    edits += 1
            left_sheet, right_sheet = grid_sheet(base), grid_sheet(right)
            result = compare_sheet("s", left_sheet, right_sheet)
            assert_consistent(self, Comparison(Workbook(Path("l"), {}), Workbook(Path("r"), {}), [result]))
            if kind == "insert" and result.structure["rows_added"] != inserted_rows:
                print("\nBASE", base, "\nRIGHT", right, "\nMAP", result.row_map, left_sheet.rows, right_sheet.rows)
            if kind == "insert":
                self.assertEqual(result.changes, {})
                self.assertEqual(result.structure, {"rows_added": inserted_rows, "rows_deleted": 0, "columns_added": inserted_cols, "columns_deleted": 0})
            elif kind == "delete":
                self.assertEqual(result.changes, {})
                self.assertEqual((result.structure["rows_added"], result.structure["columns_added"]), (0, 0))
            elif kind == "edit":
                self.assertFalse(any(result.structure.values()))
                self.assertLessEqual(len(result.changes), edits)


class CoreTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.left = Path(self.temp.name) / "left.xlsx"
        self.right = Path(self.temp.name) / "right.xlsx"

    def save(self, path, sheets):
        book = ExcelWorkbook()
        book.remove(book.active)
        for name, cells in sheets.items():
            sheet = book.create_sheet(name)
            for address, value in cells.items():
                sheet[address] = value
        book.save(path)
        book.close()

    def test_two_changes_on_first_and_fifth_sheet(self):
        sheets = {name: {"A1": "same", "B5": "before"} for name in ("main.irai", "main.proc", "main.m_status", "INSERT文", "main.slide")}
        self.save(self.left, sheets)
        for name in ("main.irai", "main.slide"):
            sheets[name]["B5"] = "after"
        self.save(self.right, sheets)
        result = compare_files(self.left, self.right)
        self.assertEqual(result.counts, {Change.DELETED: 0, Change.ADDED: 0, Change.MODIFIED: 2})
        self.assertEqual([len(s.changes) for s in result.sheets], [1, 0, 0, 0, 1])

    def test_blanks_zero_false_and_types(self):
        self.save(self.left, {"main": {"A1": "delete", "C1": 0, "D1": False, "E1": "1", "F1": "", "G1": " "}})
        self.save(self.right, {"main": {"B1": "add", "C1": False, "D1": False, "E1": 1, "G1": " "}})
        self.assertEqual(compare_files(self.left, self.right).counts, {Change.DELETED: 1, Change.ADDED: 1, Change.MODIFIED: 2})

    def test_sheet_union_and_missing_empty_sheet(self):
        self.save(self.left, {"z": {"A1": "x"}, "shared": {"A1": 10}, "empty": {}})
        self.save(self.right, {"shared": {"A1": 10}, "new": {"AA50": 5}})
        result = compare_files(self.left, self.right)
        self.assertEqual([s.name for s in result.sheets], ["z", "shared", "empty", "new"])
        self.assertIsNone(result.sheets[2].right)
        self.assertEqual(result.counts, {Change.DELETED: 1, Change.ADDED: 1, Change.MODIFIED: 0})
        self.assertEqual(result.sheets[-1].columns, 27)
        self.assertEqual(result.sheets[-1].rows, 50)

    def test_formula_refresh_and_no_retained_file_lock(self):
        self.save(self.left, {"main": {"A1": "=1+1"}})
        self.save(self.right, {"main": {"A1": "=2"}})
        initial = compare_files(self.left, self.right)
        self.assertEqual(initial.counts[Change.MODIFIED], 1)
        renamed = self.left.with_name("renamed.xlsx")
        self.left.rename(renamed)
        renamed.unlink()
        self.save(self.left, {"main": {"A1": "=2"}})
        self.assertEqual(compare_files(self.left, self.right).counts[Change.MODIFIED], 0)
        self.assertEqual(initial.left.sheets["main"].cells[(1, 1)].text, "=1+1")

    def test_errors(self):
        with self.assertRaises(WorkbookError):
            read_workbook(self.left)
        with self.assertRaises(WorkbookError):
            read_workbook(self.left.with_suffix(".xls"))
        self.left.write_text("invalid")
        with self.assertRaises(WorkbookError):
            read_workbook(self.left)

    def test_formatting_is_not_a_change(self):
        book = ExcelWorkbook()
        book.active["A1"] = 12
        book.save(self.left)
        book.active["A1"].number_format = "0.00"
        book.save(self.right)
        book.close()
        self.assertFalse(compare_files(self.left, self.right).sheets[0].changes)

    def test_provided_workbooks(self):
        directory = Path(__file__).resolve().parents[1] / "test-data"
        left, right = directory / "data_gen_01.xlsx", directory / "data_gen_02.xlsx"
        if not left.exists() or not right.exists():
            self.skipTest("提供された検証用ファイルがありません")
        assert_consistent(self, compare_files(left, right))
        result = compare_files(left, right, align=False)
        # These user-owned files can change during manual refresh testing.
        # Independently check current contents instead of freezing their counts.
        a, b = load_workbook(left), load_workbook(right)
        try:
            names = list(dict.fromkeys([*a.sheetnames, *b.sheetnames]))
            self.assertEqual([sheet.name for sheet in result.sheets], names)
            for sheet in result.sheets:
                lhs = a[sheet.name] if sheet.name in a.sheetnames else None
                rhs = b[sheet.name] if sheet.name in b.sheetnames else None
                expected = {}
                for row in range(1, max(lhs.max_row if lhs else 0, rhs.max_row if rhs else 0) + 1):
                    for col in range(1, max(lhs.max_column if lhs else 0, rhs.max_column if rhs else 0) + 1):
                        x, y = lhs.cell(row, col) if lhs else None, rhs.cell(row, col) if rhs else None
                        xv, yv = x.value if x else None, y.value if y else None
                        if xv in (None, "") and yv in (None, ""):
                            continue
                        if xv in (None, ""):
                            expected[row, col] = Change.ADDED
                        elif yv in (None, ""):
                            expected[row, col] = Change.DELETED
                        elif (xv, x.data_type) != (yv, y.data_type):
                            expected[row, col] = Change.MODIFIED
                self.assertEqual(sheet.changes, expected)
        finally:
            a.close()
            b.close()


if __name__ == "__main__":
    unittest.main()
