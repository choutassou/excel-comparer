import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from pathlib import Path
from tempfile import TemporaryDirectory
import time
import unittest
from unittest.mock import patch

from PySide6.QtCore import QItemSelection, QItemSelectionModel, Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import QApplication
from openpyxl import Workbook as ExcelWorkbook
from source.core import CellValue, Change, Sheet, Workbook, compare_workbooks
from source.gui import COLORS, MainWindow


class GuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.window = MainWindow()
        self.window.show()
        left = Workbook(Path("left.xlsx"), {"main": Sheet("main", {(1, 1): CellValue("before", "s"), (100, 40): CellValue("end", "s")}, 100, 40), "second": Sheet("second")})
        right = Workbook(Path("right.xlsx"), {"main": Sheet("main", {(1, 1): CellValue("after", "s"), (100, 40): CellValue("end", "s")}, 100, 40), "second": Sheet("second")})
        self.window.apply_comparison(compare_workbooks(left, right))
        self.app.processEvents()
        self.addCleanup(self.window.close)

    def test_display_copy_and_read_only(self):
        table = self.window.tables[0]
        index = table.model().index(0, 0)
        self.assertFalse(table.model().flags(index) & Qt.ItemFlag.ItemIsEditable)
        self.assertIsNotNone(index.data(Qt.ItemDataRole.BackgroundRole))
        table.setCurrentIndex(index)
        table.copy_selection()
        self.assertEqual(self.app.clipboard().text(), "before\r\n")
        self.assertEqual(self.window.cell_text.toPlainText(), "before")
        table.selectionModel().select(QItemSelection(index, table.model().index(1, 1)), QItemSelectionModel.SelectionFlag.ClearAndSelect)
        table.copy_selection()
        self.assertEqual(self.app.clipboard().text(), "before\t\r\n\t\r\n")

    def test_sync_navigation_and_tab_position(self):
        left, right = self.window.tables
        left.horizontalScrollBar().setValue(320)
        right.verticalScrollBar().setValue(190)
        self.assertEqual(right.horizontalScrollBar().value(), 320)
        self.assertEqual(left.verticalScrollBar().value(), 190)
        left.setColumnWidth(2, 210)
        self.assertEqual(right.columnWidth(2), 210)
        self.window.tabs.setCurrentIndex(1)
        self.window.tabs.setCurrentIndex(0)
        self.assertEqual(left.horizontalScrollBar().value(), 320)
        self.assertEqual(left.verticalScrollBar().value(), 190)
        self.window.navigate(1)
        self.assertEqual(right.currentIndex().data(), "after")
        self.assertEqual(left.currentIndex().data(), "before")

    def test_aligned_display_and_mode_toggle(self):
        def sheet(grid):
            cells = {(r, c): CellValue(v, "s") for r, row in enumerate(grid, 1) for c, v in enumerate(row, 1) if v}
            return Sheet("main", cells, len(grid), max(map(len, grid)))
        left = Workbook(Path("left.xlsx"), {"main": sheet([["id", "name"], ["1", "a"], ["2", "b"]])})
        right = Workbook(Path("right.xlsx"), {"main": sheet([["id", "memo", "name"], ["1", "m1", "a"], ["9", "m9", "z"], ["2", "m2", "B"]])})
        self.window.apply_comparison(compare_workbooks(left, right))
        self.app.processEvents()
        lhs, rhs = (table.model() for table in self.window.tables)
        # Display row 3 is the inserted row: filler on the left, green on the right.
        self.assertEqual(lhs.headerData(2, Qt.Orientation.Vertical), "")
        self.assertEqual(rhs.headerData(2, Qt.Orientation.Vertical), "3")
        self.assertEqual(lhs.headerData(1, Qt.Orientation.Horizontal), "")
        self.assertEqual(lhs.headerData(2, Qt.Orientation.Horizontal), "B")
        self.assertEqual(rhs.headerData(2, Qt.Orientation.Horizontal), "C")
        filler = lhs.index(2, 0).data(Qt.ItemDataRole.BackgroundRole)
        self.assertIsInstance(filler, QBrush)
        self.assertEqual(filler.style(), Qt.BrushStyle.BDiagPattern)
        self.assertEqual(rhs.index(2, 0).data(Qt.ItemDataRole.BackgroundRole), QColor(COLORS[Change.ADDED][0]))
        self.assertEqual((lhs.index(3, 2).data(), rhs.index(3, 2).data()), ("b", "B"))
        self.assertEqual(rhs.index(3, 2).data(Qt.ItemDataRole.BackgroundRole), QColor(COLORS[Change.MODIFIED][0]))
        self.assertEqual(self.window.tabs.tabText(0), "main  (3)")
        self.window.navigate(1)
        self.window.navigate(1)
        self.assertIn("左 —", self.window.statusBar().currentMessage())
        self.assertEqual(self.window.tables[1].currentIndex().data(), "9")
        self.window.align_box.setChecked(False)
        self.wait_for_load()
        self.assertFalse(self.window.comparison.aligned)
        self.assertEqual(self.window.tables[0].model().headerData(2, Qt.Orientation.Vertical), "3")
        self.assertGreater(self.window.comparison.counts[Change.MODIFIED], 3)

    def wait_for_load(self):
        deadline = time.monotonic() + 10
        while self.window.worker is not None and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(0.01)
        self.assertIsNone(self.window.worker)

    def test_async_refresh_and_failure(self):
        with TemporaryDirectory() as tmp:
            a, b = Path(tmp) / "a.xlsx", Path(tmp) / "b.xlsx"
            book = ExcelWorkbook()
            book.active["A1"] = "old"
            book.save(a)
            book.active["A1"] = "new"
            book.save(b)
            self.window.load_files(str(a), str(b))
            self.wait_for_load()
            self.assertEqual(self.window.comparison.counts[Change.MODIFIED], 1)
            book.save(a)
            book.close()
            self.window.refresh()
            self.wait_for_load()
            self.assertEqual(self.window.comparison.counts[Change.MODIFIED], 0)
            previous = self.window.comparison
            b.unlink()
            with patch("source.gui.QMessageBox.warning") as warning:
                self.window.refresh()
                self.wait_for_load()
                warning.assert_called_once()
            self.assertIs(self.window.comparison, previous)
            self.assertTrue(self.window.refresh_button.isEnabled())


if __name__ == "__main__":
    unittest.main()
