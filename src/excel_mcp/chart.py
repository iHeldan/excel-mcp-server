from enum import Enum
import logging
from typing import Any, Dict, Optional

from openpyxl.chart import (
    AreaChart,
    BarChart,
    LineChart,
    PieChart,
    Reference,
    ScatterChart,
    Series,
)
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.chart.axis import ChartLines
from openpyxl.utils import column_index_from_string, get_column_letter

from .cell_utils import parse_cell_range
from .exceptions import ValidationError, ChartError
from .workbook import safe_workbook

logger = logging.getLogger(__name__)

class ChartType(str, Enum):
    """Supported chart types"""
    LINE = "line"
    BAR = "bar"
    PIE = "pie"
    SCATTER = "scatter"
    AREA = "area"

class ChartStyle:
    """Chart style configuration"""
    def __init__(
        self,
        title_size: int = 14,
        title_bold: bool = True,
        axis_label_size: int = 12,
        show_legend: bool = True,
        legend_position: str = "r",
        show_data_labels: bool = True,
        grid_lines: bool = False,
        style_id: int = 2
    ):
        self.title_size = title_size
        self.title_bold = title_bold
        self.axis_label_size = axis_label_size
        self.show_legend = show_legend
        self.legend_position = legend_position
        self.show_data_labels = show_data_labels
        self.grid_lines = grid_lines
        self.style_id = style_id


def _extract_text_runs(rich_text: Any) -> Optional[str]:
    parts: list[str] = []
    for paragraph in getattr(rich_text, "p", []) or []:
        for run in getattr(paragraph, "r", []) or []:
            text = getattr(run, "t", None)
            if text:
                parts.append(text)
        for field in getattr(paragraph, "fld", []) or []:
            text = getattr(field, "t", None)
            if text:
                parts.append(text)
    return "".join(parts) or None


def _extract_title_text(title: Any) -> Optional[str]:
    if title is None:
        return None
    if isinstance(title, str):
        return title or None

    tx = getattr(title, "tx", None)
    if tx is not None:
        rich_text = getattr(tx, "rich", None)
        if rich_text is not None:
            extracted = _extract_text_runs(rich_text)
            if extracted:
                return extracted

        str_ref = getattr(tx, "strRef", None)
        if str_ref is not None and getattr(str_ref, "f", None):
            return str_ref.f

    str_ref = getattr(title, "strRef", None)
    if str_ref is not None and getattr(str_ref, "f", None):
        return str_ref.f

    value = getattr(title, "v", None)
    if value is not None:
        return str(value)

    return None


def _extract_reference_formula(data_source: Any) -> Optional[str]:
    if data_source is None:
        return None

    for attr_name in ("numRef", "strRef", "multiLvlStrRef"):
        reference = getattr(data_source, attr_name, None)
        if reference is not None and getattr(reference, "f", None):
            return reference.f

    return None


def _extract_chart_anchor(chart: Any) -> Optional[str]:
    anchor = getattr(chart, "anchor", None)
    marker = getattr(anchor, "_from", None)
    if marker is None:
        return None
    return f"{get_column_letter(marker.col + 1)}{marker.row + 1}"


def _extract_series_metadata(series: Any) -> Dict[str, Any]:
    metadata: Dict[str, Any] = {}

    title = _extract_title_text(getattr(series, "tx", None))
    if title:
        metadata["title"] = title

    categories = _extract_reference_formula(getattr(series, "cat", None))
    if categories:
        metadata["categories"] = categories

    x_values = _extract_reference_formula(getattr(series, "xVal", None))
    if x_values:
        metadata["x_values"] = x_values

    values = _extract_reference_formula(getattr(series, "val", None))
    if values:
        metadata["values"] = values

    y_values = _extract_reference_formula(getattr(series, "yVal", None))
    if y_values:
        metadata["y_values"] = y_values

    return metadata


def _chart_type_name(chart: Any) -> str:
    class_name = type(chart).__name__
    if class_name.endswith("Chart"):
        return class_name.removesuffix("Chart").lower()
    return class_name.lower()


def list_charts(
    filepath: str,
    sheet_name: Optional[str] = None,
) -> list[dict[str, Any]]:
    """List embedded charts for one worksheet or the whole workbook."""
    try:
        with safe_workbook(filepath) as wb:
            if sheet_name is not None and sheet_name not in wb.sheetnames:
                raise ValidationError(f"Sheet '{sheet_name}' not found")

            sheet_names = [sheet_name] if sheet_name is not None else list(wb.sheetnames)
            charts: list[dict[str, Any]] = []

            for current_sheet_name in sheet_names:
                worksheet = wb[current_sheet_name]
                for chart_index, chart in enumerate(getattr(worksheet, "_charts", []), start=1):
                    series = getattr(chart, "ser", None) or list(getattr(chart, "series", []))
                    chart_info = {
                        "sheet_name": current_sheet_name,
                        "chart_index": chart_index,
                        "chart_type": _chart_type_name(chart),
                        "anchor": _extract_chart_anchor(chart),
                        "title": _extract_title_text(getattr(chart, "title", None)),
                        "x_axis_title": _extract_title_text(
                            getattr(getattr(chart, "x_axis", None), "title", None)
                        ),
                        "y_axis_title": _extract_title_text(
                            getattr(getattr(chart, "y_axis", None), "title", None)
                        ),
                        "legend_position": getattr(getattr(chart, "legend", None), "position", None),
                        "style": getattr(chart, "style", None),
                        "series": [_extract_series_metadata(item) for item in series],
                    }
                    charts.append({key: value for key, value in chart_info.items() if value is not None})

            return charts

    except ValidationError:
        raise
    except Exception as e:
        logger.error(f"Failed to list charts: {e}")
        raise ChartError(str(e))


def _validate_target_cell(target_cell: str) -> None:
    if not target_cell:
        raise ValidationError("Invalid target cell format: target cell is required")

    column_part = "".join(character for character in target_cell if character.isalpha())
    row_part = "".join(character for character in target_cell if character.isdigit())
    if not column_part or not row_part:
        raise ValidationError(f"Invalid target cell format: {target_cell}")

    try:
        column_index_from_string(column_part)
        row_index = int(row_part)
    except ValueError as e:
        raise ValidationError(f"Invalid target cell: {str(e)}") from e

    if row_index < 1:
        raise ValidationError(f"Invalid target cell: {target_cell}")


def create_chart_in_sheet(
    filepath: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = "",
    style: Optional[Dict] = None
) -> dict[str, Any]:
    """Create chart in sheet with enhanced styling options"""
    # Ensure style dict exists and defaults to showing data labels
    if style is None:
        style = {"show_data_labels": True}
    else:
        # If caller omitted the flag, default to True
        style.setdefault("show_data_labels", True)
    try:
        with safe_workbook(filepath, save=True) as wb:
            if sheet_name not in wb.sheetnames:
                logger.error(f"Sheet '{sheet_name}' not found")
                raise ValidationError(f"Sheet '{sheet_name}' not found")

            worksheet = wb[sheet_name]

            # Parse the data range
            if "!" in data_range:
                range_sheet_name, cell_range = data_range.split("!")
                if range_sheet_name not in wb.sheetnames:
                    logger.error(f"Sheet '{range_sheet_name}' referenced in data range not found")
                    raise ValidationError(f"Sheet '{range_sheet_name}' referenced in data range not found")
            else:
                cell_range = data_range

            try:
                start_cell, end_cell = cell_range.split(":")
                start_row, start_col, end_row, end_col = parse_cell_range(start_cell, end_cell)
            except ValueError as e:
                logger.error(f"Invalid data range format: {e}")
                raise ValidationError(f"Invalid data range format: {str(e)}")

            # Validate chart type
            chart_classes = {
                "line": LineChart,
                "bar": BarChart,
                "pie": PieChart,
                "scatter": ScatterChart,
                "area": AreaChart
            }

            chart_type_lower = chart_type.lower()
            ChartClass = chart_classes.get(chart_type_lower)
            if not ChartClass:
                logger.error(f"Unsupported chart type: {chart_type}")
                raise ValidationError(
                    f"Unsupported chart type: {chart_type}. "
                    f"Supported types: {', '.join(chart_classes.keys())}"
                )

            chart = ChartClass()

            # Basic chart settings
            chart.title = title
            if hasattr(chart, "x_axis"):
                chart.x_axis.title = x_axis
            if hasattr(chart, "y_axis"):
                chart.y_axis.title = y_axis

            try:
                # Create data references
                if chart_type_lower == "scatter":
                    # For scatter charts, create series for each pair of columns
                    for col in range(start_col + 1, end_col + 1):
                        x_values = Reference(
                            worksheet,
                            min_row=start_row + 1,
                            max_row=end_row,
                            min_col=start_col
                        )
                        y_values = Reference(
                            worksheet,
                            min_row=start_row + 1,
                            max_row=end_row,
                            min_col=col
                        )
                        series = Series(y_values, x_values, title_from_data=True)
                        chart.series.append(series)
                else:
                    # For other chart types
                    data = Reference(
                        worksheet,
                        min_row=start_row,
                        max_row=end_row,
                        min_col=start_col + 1,
                        max_col=end_col
                    )
                    cats = Reference(
                        worksheet,
                        min_row=start_row + 1,
                        max_row=end_row,
                        min_col=start_col
                    )
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(cats)
            except Exception as e:
                logger.error(f"Failed to create chart data references: {e}")
                raise ChartError(f"Failed to create chart data references: {str(e)}")

            # Apply style if provided
            try:
                if style.get("show_legend", True):
                    chart.legend = Legend()
                    chart.legend.position = style.get("legend_position", "r")
                else:
                    chart.legend = None

                if style.get("show_data_labels", False):
                    data_labels = DataLabelList()
                    # Gather optional overrides
                    dlo = style.get("data_label_options", {}) if isinstance(style.get("data_label_options", {}), dict) else {}

                    # Helper to read bool with fallback
                    def _opt(name: str, default: bool) -> bool:
                        return bool(dlo.get(name, default))

                    # Apply options -- Excel will concatenate any that are set to True
                    data_labels.showVal = _opt("show_val", True)
                    data_labels.showCatName = _opt("show_cat_name", False)
                    data_labels.showSerName = _opt("show_ser_name", False)
                    data_labels.showLegendKey = _opt("show_legend_key", False)
                    data_labels.showPercent = _opt("show_percent", False)
                    data_labels.showBubbleSize = _opt("show_bubble_size", False)

                    chart.dataLabels = data_labels

                if style.get("grid_lines", False):
                    if hasattr(chart, "x_axis"):
                        chart.x_axis.majorGridlines = ChartLines()
                    if hasattr(chart, "y_axis"):
                        chart.y_axis.majorGridlines = ChartLines()
            except Exception as e:
                logger.error(f"Failed to apply chart style: {e}")
                raise ChartError(f"Failed to apply chart style: {str(e)}")

            # Set chart size
            chart.width = 15
            chart.height = 7.5

            # Create drawing and anchor
            try:
                _validate_target_cell(target_cell)
                worksheet.add_chart(chart, target_cell)
            except ValidationError:
                raise
            except Exception as e:
                logger.error(f"Failed to create chart drawing: {e}")
                raise ChartError(f"Failed to create chart drawing: {str(e)}")

        return {
            "message": f"{chart_type.capitalize()} chart created successfully",
            "details": {
                "type": chart_type,
                "location": target_cell,
                "data_range": data_range
            }
        }

    except (ValidationError, ChartError):
        raise
    except Exception as e:
        logger.error(f"Unexpected error creating chart: {e}")
        raise ChartError(f"Unexpected error creating chart: {str(e)}")
