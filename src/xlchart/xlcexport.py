import re
import sys
from os import PathLike
from pathlib import Path

import win32com.client.gencache


def usage():
    cmd = Path(__file__).name
    print(f"Usage: {cmd} <workbook> [dest_dir]")
    print(f"       {cmd} <directory> [dest_dir]")


def main():

    if len(sys.argv) < 2:
        usage()
        exit()

    target_path = Path(sys.argv[1]).resolve()
    if len(sys.argv) > 2:
        dest_path = Path(sys.argv[2])
    elif target_path.is_file():
        dest_path = target_path.parent
    else:
        dest_path = target_path

    if target_path.is_file():
        export(target_path, dest_path)
        return

    if target_path.is_dir():
        for target_book in target_path.glob("*.xlsx"):
            print(target_book, file=sys.stderr)
            export(target_book, dest_path)
        return

    print(f"Error: No such file or directory: {target_path}", file=sys.stderr)


def export(workbook_path: str | PathLike, dest_path: str | PathLike):

    dest_dir = Path(dest_path)
    if dest_dir.exists() and not dest_dir.is_dir():
        print(f"Error: Not a directory: {dest_path}", file=sys.stderr)
        return

    xl = win32com.client.gencache.EnsureDispatch("Excel.Application")
    wb = xl.Workbooks.Open(Path(workbook_path).resolve())

    # ファイル名と同じ名前のディレクトリに出力する
    dest_dir = dest_dir.joinpath(Path(workbook_path).stem)
    if not dest_dir.exists():
        dest_dir.mkdir()

    # 埋め込みグラフ
    for sheet in wb.Worksheets:
        for obj in sheet.ChartObjects():
            name = escape_name(f"{sheet.Name}_{obj.Name}")
            dest_file = dest_dir.joinpath(f"{name}.png")
            if dest_file.exists():
                print(f"Error: File already exists: {str(dest_file)}", file=sys.stderr)
                continue
            obj.Chart.Export(dest_file)

    # グラフシート
    for chart in wb.Charts:
        name = escape_name(chart.Name)
        dest_file = dest_dir.joinpath(f"{name}.png")
        if dest_file.exists():
            print(f"Error: File already exists: {str(dest_file)}", file=sys.stderr)
            continue
        chart.Export(dest_file)

    wb.Close(SaveChanges=False)
    xl.Quit()


def escape_name(name: str) -> str:

    # 全角を半角に変換
    table = str.maketrans({chr(0xFF01 + i): chr(0x21 + i) for i in range(94)})
    name = name.translate(table)

    # 空白文字をアンダースコアに変換
    name = re.sub(r"\s|　", "_", name)

    # ファイル名に使用できない文字を削除
    name = re.sub(r'[\\/:*?"<>|]', "", name)

    return name


if __name__ == "__main__":
    main()
