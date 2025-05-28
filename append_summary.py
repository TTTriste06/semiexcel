from openpyxl.styles import PatternFill

def standardize(val):
    if val is None:
        return ""
    return str(val).strip().replace("\u3000", " ").strip('\'"“”‘’')


def append_forecast_unmatched_to_summary_by_keys(ws_summary, ws_forecast, unmatched_names, name_col=2, start_col_forecast=4):
    """
    将预测表中品名在 unmatched_names 列表中的行，追加到汇总表末尾，并复制预测信息。

    参数：
    - ws_summary: openpyxl 的汇总 worksheet
    - ws_forecast: openpyxl 的预测 worksheet
    - unmatched_names: list[str]，未匹配的品名（如 keys_main）
    - name_col: 品名所在列号（默认第3列）
    - start_col_forecast: 从第几列开始是预测值（默认第4列）
    """
    unmatched_set = set(standardize(name) for name in unmatched_names)
    max_summary_row = ws_summary.max_row

    for row in range(2, ws_forecast.max_row + 1):
        name = standardize(ws_forecast.cell(row=row, column=name_col).value)
        if name in unmatched_set:
            new_row = max_summary_row + 1
            # 填品名
            ws_summary.cell(row=new_row, column=name_col).value = name

            # 拷贝预测数据
            for col in range(start_col_forecast, ws_forecast.max_column + 1):
                value = ws_forecast.cell(row=row, column=col).value
                ws_summary.cell(row=new_row, column=col).value = value

            max_summary_row += 1
