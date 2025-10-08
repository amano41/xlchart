import json
import sys
from os import PathLike
from pathlib import Path

import win32com.client.gencache

from . import xlcparse


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
        xl = None
        try:
            xl = win32com.client.gencache.EnsureDispatch("Excel.Application")
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
                xl.Quit()
                del xl
        return

    print(f"Error: No such file or directory: {target_path}", file=sys.stderr)
    sys.exit(1)


def dump(workbook_path: str | PathLike) -> dict:
    xl = None
    try:
        xl = win32com.client.gencache.EnsureDispatch("Excel.Application")
        return _dump(xl, workbook_path)
    except Exception:
        raise
    finally:
        if xl is not None:
            xl.Quit()
            del xl


def _dump(xl, workbook_path: str | PathLike) -> dict:
    wb = None
    try:
        wb = xl.Workbooks.Open(Path(workbook_path).resolve(), ReadOnly=True, UpdateLinks=False)
        if wb is None:
            raise RuntimeError(f"Failed to open workbook: {workbook_path}")
        data = xlcparse.parse_book(wb)
    except Exception:
        raise
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
            del wb
    return data


if __name__ == "__main__":
    main()
