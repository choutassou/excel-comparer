import argparse
import sys

from PySide6.QtWidgets import QApplication

from source.gui import MainWindow


def main():
    parser = argparse.ArgumentParser(description="Excelファイルの左右比較")
    parser.add_argument("left", nargs="?")
    parser.add_argument("right", nargs="?")
    args = parser.parse_args()
    if bool(args.left) != bool(args.right):
        parser.error("比較する左右2つのファイルを指定してください。")
    app = QApplication(sys.argv[:1])
    app.setApplicationName("Excel比較")
    window = MainWindow()
    window.show()
    if args.left:
        window.load_files(args.left, args.right)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
