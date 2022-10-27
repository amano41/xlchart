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
        data = dump(target_path)
        print(json.dumps(data, indent=4, ensure_ascii=False))
        return

    if target_path.is_dir():
        for target_book in target_path.glob("*.xlsx"):
            print(target_book, file=sys.stderr)
            data = dump(target_book)
            output = target_book.with_suffix(".json")
            with output.open("w", encoding="utf-8", newline="\n") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
                f.write("\n")
        return

    print(f"Error: No such file or directory: {target_path}", file=sys.stderr)


def dump(workbook_path: str | PathLike) -> dict:
    xl = win32com.client.gencache.EnsureDispatch("Excel.Application")
    wb = xl.Workbooks.Open(Path(workbook_path).resolve())
    try:
        data = xlcparse.parse_book(wb)
    except Exception:
        raise
    finally:
        wb.Close(SaveChanges=False)
        xl.Quit()
    return data


if __name__ == "__main__":
    main()
