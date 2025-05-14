import re

def extract_months_from_columns(columns):
    """
    从列名中提取所有符合 `_YYYY-MM` 格式的月份。

    参数:
    - columns: DataFrame.columns
    返回:
    - 排序后的月份字符串列表 ['2024-01', '2024-02', ...]
    """
    month_pattern = re.compile(r"_(\d{4}-\d{2})$")
    months = set()
    for col in columns:
        match = month_pattern.search(col)
        if match:
            months.add(match.group(1))
    return sorted(months)


def process_history_columns(pivoted, config, selected_month):
    """
    将小于等于 selected_month 的列合并为“历史订单数量”和“历史未交订单数量”。

    参数:
    - pivoted: 已透视的 DataFrame
    - config: 当前表的 pivot_config（需含 index 字段）
    - selected_month: 用户选择的截止月份（YYYY-MM 字符串）
    
    返回:
    - 处理后的 pivoted DataFrame
    """
    if not selected_month:
        return pivoted

    # 找出所有小于等于选定月份的列（如订单数量_2023-12）
    history_cols = [
        col for col in pivoted.columns
        if '_' in col and col.split('_')[-1][:4].isdigit() and col.split('_')[-1] <= selected_month
    ]

    history_order_cols = [col for col in history_cols if '订单数量' in col and '未交订单数量' not in col]
    history_pending_cols = [col for col in history_cols if '未交订单数量' in col]

    # 合并为两列
    if history_order_cols:
        pivoted['历史订单数量'] = pivoted[history_order_cols].sum(axis=1)
    if history_pending_cols:
        pivoted['历史未交订单数量'] = pivoted[history_pending_cols].sum(axis=1)

    # 删除原始月列
    pivoted.drop(columns=history_cols, inplace=True)

    # 排列列顺序
    fixed_cols = [col for col in pivoted.columns if col not in ['历史订单数量', '历史未交订单数量']]
    if '历史订单数量' in pivoted.columns:
        fixed_cols.insert(len(config['index']), '历史订单数量')
    if '历史未交订单数量' in pivoted.columns:
        fixed_cols.insert(len(config['index']) + 1, '历史未交订单数量')

    return pivoted[fixed_cols]
