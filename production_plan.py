import re
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter
from dateutil.relativedelta import relativedelta
from datetime import datetime
from excel_utils import adjust_column_width


def add_colored_monthly_plan_headers(ws, start_col: int, start_date: datetime, pivot_unfulfilled) -> int:
    """
    向“产品生产计划” Sheet 添加多月字段组表头（合并单元格，字段名，彩色背景）。
    
    参数：
    - ws: openpyxl Worksheet 对象
    - start_col: 从第几列开始插入月份组（通常是已有字段后的第1列）
    - start_date: 用户在界面中选择的排产起始月份（datetime 对象）
    - pivot_unfulfilled: 未交订单透视表，用于提取最大月份字段
    
    返回：
    - 最终写入结束的列号（用于后续插入数据）
    """

    # ✅ 每月字段列名
    monthly_fields = [
        "成品投单计划", "投单计划调整", "半成品投单计划",
        "成品可行投单", "半成品可行投单", "成品实际投单",
        "回货计划", "回货实际"
    ]

    # ✅ 每月对应背景色（12个以内自动轮换）
    month_colors = [
        "FFFF00", "32CD32", "9932CC", "FFB6C1", "FFA500", "87CEFA"
    ]

    # ✅ 提取最大月份（从未交订单列名中解析）
    month_pattern = re.compile(r"(\d{4})年(\d{1,2})月.*未交订单数量")
    max_month = None
    for col in pivot_unfulfilled.columns:
        match = month_pattern.match(col)
        if match:
            year, month = int(match.group(1)), int(match.group(2))
            dt = datetime(year, month, 1)
            if not max_month or dt > max_month:
                max_month = dt
    end_date = max_month if max_month else start_date + relativedelta(months=6)

    # ✅ 开始写入月份组表头
    current_col = start_col
    month_index = 0
    while start_date <= end_date:
        fill_color = PatternFill("solid", fgColor=month_colors[month_index % len(month_colors)])
        month_name = start_date.strftime("%-m月")

        # 第一行：合并单元格并写月份名
        ws.merge_cells(
            start_row=1, start_column=current_col,
            end_row=1, end_column=current_col + len(monthly_fields) - 1
        )
        cell = ws.cell(row=1, column=current_col)
        cell.value = month_name
        cell.fill = fill_color
        cell.alignment = Alignment(horizontal="center", vertical="center")

        # 第二行：写字段名
        for offset, field in enumerate(monthly_fields):
            col = current_col + offset
            ws.cell(row=2, column=col, value=field)
            ws.cell(row=2, column=col).fill = fill_color

        # 下一月
        current_col += len(monthly_fields)
        start_date += relativedelta(months=1)
        month_index += 1

    # ✅ 自动调整新列区域的列宽
    adjust_column_width(ws, ws.title)

    return current_col  # 返回最后写入的列号
