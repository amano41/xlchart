import re
from typing import Optional

from win32com.client import constants


def parse_book(book) -> dict:
    data = dict()
    # 埋め込みグラフ
    for sheet in book.Worksheets:
        data.update(parse_sheet(sheet))
    # グラフシート
    for chart in book.Charts:
        data[chart.Name] = parse_chart(chart)
    return data


def parse_sheet(sheet) -> dict:
    data = dict()
    # 埋め込みグラフの名前は ChartObject から取得する
    for obj in sheet.ChartObjects():
        data[obj.Name] = parse_chart(obj.Chart, obj.Name)
    return data


def parse_chart(chart, name: Optional[str] = None) -> dict:

    data = dict()

    if name is None:
        name = chart.Name

    data["name"] = name
    data["chart-type"] = chart.ChartType

    if chart.HasTitle:
        data["title"] = chart.ChartTitle.Text
        data["title-overlay"] = chart.ChartTitle.IncludeInLayout
    else:
        data["title"] = ""
        data["title-overlay"] = 0

    if chart.HasLegend:
        data["legend-position"] = chart.Legend.Position
    else:
        data["legend-position"] = 0

    data["axis"] = list()
    for axis in chart.Axes():
        data["axis"].append(parse_axis(axis, chart.Type))

    series = list()
    for i, group in enumerate(chart.ChartGroups()):
        series.extend(parse_series_by_group(group, i + 1))
    data["series"] = sorted(series, key=lambda d: d["index"])

    return data


def parse_axis(axis, chart_type: str):

    data = dict()

    data["axis-type"] = axis.Type
    data["axis-group"] = axis.AxisGroup

    if axis.HasTitle:
        data["title"] = axis.AxisTitle.Caption
        data["title-orientation"] = axis.AxisTitle.Orientation

    # 散布図の X 軸は数値軸だが Type は xlCategory になっている
    if is_value_axis(axis) or is_scatter_chart(chart_type):
        data["min-scale"] = axis.MinimumScale
        data["min-scale-auto"] = axis.MinimumScaleIsAuto
        data["max-scale"] = axis.MaximumScale
        data["max-scale-auto"] = axis.MaximumScaleIsAuto
        data["major-unit"] = axis.MajorUnit
        data["major-unit-auto"] = axis.MajorUnitIsAuto
        data["minor-unit"] = axis.MinorUnit
        data["minor-unit-auto"] = axis.MinorUnitIsAuto
    else:
        if is_category_axis(axis):
            data["category-names"] = axis.CategoryNames
        data["tick-label-spacing"] = axis.TickLabelSpacing
        data["tick-label-spacing-auto"] = axis.TickLabelSpacingIsAuto

    data["tick-label-format"] = axis.TickLabels.NumberFormatLocal

    # レーダーチャートでは軸のオプションは利用できない
    if is_radar_chart(chart_type):
        return data

    # 系列軸では軸の交点を指定できない
    if not is_series_axis(axis):
        data["crosses"] = axis.Crosses
        data["crosses-at"] = axis.CrossesAt

    data["reverse"] = axis.ReversePlotOrder

    return data


def parse_series_by_group(group, group_number: int = 1):

    series = list()

    for s in group.SeriesCollection():
        data = parse_series(s)
        data["overlap"] = group.Overlap
        data["gap-width"] = group.GapWidth
        data["chart-group"] = group_number
        series.append(data)

    return series


def parse_series(series):

    data = dict()

    # Formula の要素にカンマが含まれている場合の対策
    formula = re.sub(r"=SERIES\((.+)\)", r"\1", series.Formula)
    formula = re.sub(r"\([^\)]+\)", lambda m: m.group().replace(",", "\t"), formula)
    formula = re.sub(r"{[^}]+}", lambda m: m.group().replace(",", "\t"), formula)

    s_name, x_vals, y_vals, index = formula.split(",")

    s_name = re.sub(r"^\"(.+)\"$", r"\1", s_name)
    s_name = s_name.replace("\t", ",")
    x_vals = x_vals.replace("\t", ",")
    y_vals = y_vals.replace("\t", ",")

    data["index"] = int(index) - 1
    data["name"] = series.Name
    data["type"] = series.ChartType

    data["formula"] = series.Formula
    data["data-range-name"] = s_name
    data["data-range-x-values"] = x_vals
    data["data-range-y-values"] = y_vals

    if series.HasDataLabels:
        labels = series.DataLabels()
        data["data-labels-range"] = labels.ShowRange
        data["data-labels-name"] = labels.ShowSeriesName
        data["data-labels-x-values"] = labels.ShowCategoryName
        data["data-labels-y-values"] = labels.ShowValue
        data["data-labels-marker"] = labels.ShowLegendKey
        data["leader-lines"] = series.HasLeaderLines

    data["axis-group"] = series.AxisGroup

    return data


def is_scatter_chart(chart_type: str) -> bool:
    return chart_type in (
        # fmt: off
        constants.xlXYScatter,
        constants.xlXYScatterLines,
        constants.xlXYScatterLinesNoMarkers,
        constants.xlXYScatterSmooth,
        constants.xlXYScatterSmoothNoMarkers,
        # fmt: on
    )


def is_radar_chart(chart_type: str) -> bool:
    return chart_type in (
        # fmt: off
        constants.xlRadar,
        constants.xlRadarFilled,
        constants.xlRadarMarkers,
        # fmt: on
    )


def is_value_axis(axis) -> bool:
    return axis.Type == constants.xlValue


def is_category_axis(axis) -> bool:
    return axis.Type == constants.xlCategory


def is_series_axis(axis) -> bool:
    return axis.Type == constants.xlSeriesAxis
