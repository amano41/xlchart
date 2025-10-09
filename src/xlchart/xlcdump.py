import json
import sys
from os import PathLike
from pathlib import Path

from . import xlcparse
from ._xlapp import _init_excel, _new_excel, _quit_excel


def usage():
    cmd = Path(__file__).name
    print(f"Usage: {cmd} <workbook>")
    print(f"       {cmd} <directory>")


def main():

    if len(sys.argv) != 2:
        usage()
        exit()

    target_path = Path(sys.argv[1]).resolve()

    if target_path.is_file():
        try:
            data = dump(target_path)
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)
        print(json.dumps(data, indent=4, ensure_ascii=False))
        return

    if target_path.is_dir():
        _init_excel()
        xl = None
        try:
            xl = _new_excel()
            for target_book in target_path.glob("*.xlsx"):
                print(target_book, file=sys.stderr)
                try:
                    data = _dump(xl, target_book)
                except Exception as e:
                    print(f"Error: {e}", file=sys.stderr)
                    continue
                output = target_book.with_suffix(".json")
                with output.open("w", encoding="utf-8", newline="\n") as f:
                    json.dump(data, f, indent=4, ensure_ascii=False)
                    f.write("\n")
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            exit(1)
        finally:
            if xl is not None:
                _quit_excel(xl)
        return

    print(f"Error: No such file or directory: {target_path}", file=sys.stderr)
    sys.exit(1)


def dump(workbook_path: str | PathLike) -> dict:
    _init_excel()
    xl = None
    try:
        xl = _new_excel()
        return _dump(xl, workbook_path)
    finally:
        if xl is not None:
            _quit_excel(xl)


def _dump(xl, workbook_path: str | PathLike) -> dict:
    wb = None
    try:
        wb = xl.Workbooks.Open(Path(workbook_path).resolve(), ReadOnly=True, UpdateLinks=False)
        if wb is None:
            raise RuntimeError(f"Failed to open workbook: {workbook_path}")
        data = xlcparse.parse_book(wb)
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
            del wb
    return data


if __name__ == "__main__":
    main()
