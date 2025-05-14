import pandas as pd

def merge_safety_inventory(summary_df, safety_df):
    """
    将安全库存表中 Wafer 和 Part 信息合并到汇总数据中。
    
    参数:
    - summary_df: 汇总后的未交订单表，包含 '晶圆品名'、'规格'、'品名'
    - safety_df: 安全库存表，包含 'WaferID', 'OrderInformation', 'ProductionNO.', ' InvWaf', ' InvPart'
    
    返回:
    - 合并后的汇总 DataFrame，增加了 ' InvWaf' 和 ' InvPart' 两列
    """

    # 重命名列用于匹配
    safety_df = safety_df.rename(columns={
        'WaferID': '晶圆品名',
        'OrderInformation': '规格',
        'ProductionNO.': '品名'
    }).copy()

    # 添加标记列（可选，用于调试或统计）
    safety_df['已匹配'] = False

    # 合并：left join 确保 summary_df 保留所有行
    merged = summary_df.merge(
        safety_df[['晶圆品名', '规格', '品名', ' InvWaf', ' InvPart']],
        on=['晶圆品名', '规格', '品名'],
        how='left'
    )

    return merged
