from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font
from openpyxl.styles import PatternFill

def adjust_column_width(writer, sheet_name, df):
    """
    自动调整 Excel 工作表中各列的宽度以适应内容长度。

    参数:
    - writer: pandas 的 ExcelWriter 对象
    - sheet_name: 要调整的工作表名称
    - df: 对应写入工作表的 DataFrame 数据
    """
    worksheet = writer.sheets[sheet_name]
    for idx, col in enumerate(df.columns, 1):
        # 获取该列中所有字符串长度的最大值
        max_content_len = df[col].astype(str).str.len().max()
        header_len = len(str(col))
        column_width = max(max_content_len, header_len) * 1.2 + 5
        worksheet.column_dimensions[get_column_letter(idx)].width = min(column_width, 50)

def merge_header_for_summary(ws, df, label_ranges):
    """
    给指定汇总列添加顶部合并行标题（如“安全库存”“未交订单”）

    参数:
    - ws: openpyxl worksheet
    - df: summary DataFrame
    - label_ranges: dict，键是标题文字，值是列名范围元组，如：
        {
            "安全库存": (" InvWaf", " InvPart"),
            "未交订单": ("总未交订单", "未交订单数量_2025-08")
        }
    """

    # 插入一行作为新表头（原表头往下挪）
    ws.insert_rows(1)
    header_row = list(df.columns)

    for label, (start_col_name, end_col_name) in label_ranges.items():
        if start_col_name not in header_row or end_col_name not in header_row:
            continue

        start_idx = header_row.index(start_col_name) + 1  # Excel index starts from 1
        end_idx = header_row.index(end_col_name) + 1

        col_letter_start = get_column_letter(start_idx)
        col_letter_end = get_column_letter(end_idx)

        merge_range = f"{col_letter_start}1:{col_letter_end}1"
        ws.merge_cells(merge_range)
        cell = ws[f"{col_letter_start}1"]
        cell.value = label
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(bold=True)

def highlight_unused_safety_rows(ws, safety_df, unused_rows_df):
    """
    将未被匹配的安全库存行在 Excel 中标红。

    参数:
    - ws: openpyxl 的 worksheet 对象
    - safety_df: 完整的安全库存 DataFrame（用于定位行号）
    - unused_rows_df: 未匹配到的行，需标红
    """

    # 设置红色填充
    red_fill = PatternFill(start_color='FFFF6666', end_color='FFFF6666', fill_type='solid')

    # 识别需要标红的行索引
    key_cols = ['晶圆品名', '规格', '品名']
    unused_set = set(
        tuple(row) for row in unused_rows_df[key_cols].itertuples(index=False, name=None)
    )

    # 遍历 worksheet 并标红匹配行（跳过 header）
    for row_idx in range(2, ws.max_row + 1):
        wafer = ws.cell(row=row_idx, column=1).value
        spec = ws.cell(row=row_idx, column=2).value
        name = ws.cell(row=row_idx, column=3).value
        if (wafer, spec, name) in unused_set:
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_idx).fill = red_fill

