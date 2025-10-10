import json
import sys
from os import PathLike
from pathlib import Path
from typing import Final, Sequence

import tomli
import win32com.client.gencache

from . import xlcparse

AXIS: Final[dict[int, str]] = {1: "x-axis", 2: "y-axis", 3: "series-axis"}

RESULT_TYPE = tuple[str, str, str, bool]


def usage():
    cmd = Path(__file__).name
    print(f"Usage: {cmd} <workbook> <answer>")
    print(f"       {cmd} <directory> <answer>")


def main():

    if len(sys.argv) != 3:
        usage()
        exit()

    target_path = Path(sys.argv[1])
    answer_path = Path(sys.argv[2])

    # 採点基準
    try:
        answer = load_answer(answer_path)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return

    # 採点対象がファイルの場合は標準出力に出力
    if target_path.is_file():
        try:
            for r in check_file(target_path, answer):
                print("\t".join(map(str, r)))
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            return
        return

    # 採点対象がディレクトリの場合はファイルごとに結果を保存
    if target_path.is_dir():
        for target_book in target_path.glob("*.xlsx"):
            print(target_book, file=sys.stderr)
            try:
                result = check_file(target_book, answer)
                output = target_book.with_suffix(".tsv")
                with output.open("w", encoding="utf-8", newline="\n") as f:
                    f.write("\t".join(("Chart", "Property", "Value", "Result")))
                    for r in result:
                        f.write("\t".join(map(str, r)) + "\n")
            except Exception as e:
                print(f"Error: {e}", file=sys.stderr)
                continue
        return

    # 読み込めなかった場合はエラー
    print(f"Error: No such file or directory: {target_path}", file=sys.stderr)


def load_answer(file_path: str | PathLike) -> dict:
    p = Path(file_path)
    if p.suffix == ".toml":
        with p.open("rb") as f:
            data = tomli.load(f)
    elif p.suffix == ".json":
        with p.open("r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        raise ValueError(f"Unsupported file type: {file_path}")
    return data


def load_target(file_path: str | PathLike) -> dict:
    xl = win32com.client.gencache.EnsureDispatch("Excel.Application")
    wb = xl.Workbooks.Open(Path(file_path).resolve())
    try:
        data = xlcparse.parse_book(wb)
    except Exception:
        raise
    finally:
        wb.Close(SaveChanges=False)
        xl.Quit()
    return data


def check_file(workbook_path: str | PathLike, answer: dict) -> list[RESULT_TYPE]:
    target = load_target(workbook_path)
    return check(target, answer)


def check(target: dict, answer: dict) -> list[RESULT_TYPE]:

    result = []

    # グラフごとにチェック
    for chart_name in answer.keys():

        answer_chart = answer.get(chart_name, {})
        target_chart = target.get(chart_name, {})

        # プロパティごとにチェック
        for prop_name in answer_chart.keys():

            # Axis
            if prop_name == "axis":
                target_list = target_chart.get("axis", [])
                answer_list = answer_chart.get("axis", [])
                result.extend(check_axis(target_list, answer_list, chart_name))
                continue

            # Series
            if prop_name == "series":
                target_list = target_chart.get("series", [])
                answer_list = answer_chart.get("series", [])
                result.extend(check_series(target_list, answer_list, chart_name))
                continue

            if prop_name == "bins":
                target_list = target_chart.get("bins", [])
                answer_list = answer_chart.get("bins", [])
                result.extend(check_bins(target_list, answer_list, chart_name))
                continue

            # その他
            target_value = target_chart.get(prop_name, "")
            answer_value = answer_chart.get(prop_name, "")
            if isinstance(answer_value, Sequence):
                correct = list(target_value) == list(answer_value)
            else:
                correct = target_value == answer_value
            result.append((chart_name, prop_name, target_value, correct))

    return result


def check_axis(target_list: list[dict], answer_list: list[dict], chart_name: str) -> list[RESULT_TYPE]:

    result = []

    for answer in answer_list:

        axis_type = answer.get("axis-type", 1)
        axis_group = answer.get("axis-group", 1)

        # type と group が一致する target を探す
        for item in target_list:
            t = item.get("axis-type", 1)
            g = item.get("axis-group", 1)
            if t == axis_type and g == axis_group:
                target = item
                break
        else:
            target = {}

        # answer で指定されたプロパティについてチェック
        for prop_name in answer:
            if prop_name in ("axis-type", "axis-group"):
                continue
            label = f"{AXIS[axis_type]}{axis_group}.{prop_name}"
            answer_value = answer.get(prop_name, "")
            target_value = target.get(prop_name, "")
            if isinstance(answer_value, Sequence):
                correct = list(target_value) == list(answer_value)
            else:
                correct = target_value == answer_value
            result.append((chart_name, label, target_value, correct))

    return result


def check_series(target_list: list[dict], answer_list: list[dict], chart_name: str) -> list[RESULT_TYPE]:

    result = []

    for i, answer in enumerate(answer_list):

        index = answer.get("index", i)

        # index が一致する target を探す
        for item in target_list:
            if item.get("index", -1) == index:
                target = item
                break
        else:
            target = {}

        # answer で指定されたプロパティについてチェック
        for prop_name in answer:
            if prop_name == "index":
                continue
            label = f"series{index}.{prop_name}"
            answer_value = answer.get(prop_name, "")
            target_value = target.get(prop_name, "")
            if isinstance(answer_value, Sequence):
                correct = list(target_value) == list(answer_value)
            else:
                correct = target_value == answer_value
            result.append((chart_name, label, target_value, correct))

    return result


def check_bins(target_list: list[dict], answer_list: list[dict], chart_name: str) -> list[RESULT_TYPE]:

    result = []

    for i, answer in enumerate(answer_list):

        chart_group = answer.get("chart-group", i + 1)

        # chart-group が一致する target を探す
        for item in target_list:
            if item.get("chart-group", -1) == chart_group:
                target = item
                break
        else:
            target = {}

        # answer で指定されたプロパティについてチェック
        for prop_name in answer:
            if prop_name == "chart-group":
                continue
            label = f"bins{chart_group}.{prop_name}"
            answer_value = answer.get(prop_name, "")
            target_value = target.get(prop_name, "")
            if isinstance(answer_value, Sequence):
                correct = list(target_value) == list(answer_value)
            else:
                correct = target_value == answer_value
            result.append((chart_name, label, target_value, correct))

    return result


if __name__ == "__main__":
    main()
