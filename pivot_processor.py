import os
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from config import CONFIG
from excel_utils import adjust_column_width
from mapping_utils import apply_mapping_and_merge

FIELD_MAPPINGS = {
    "赛卓-未交订单": {"规格": "规格", "品名": "品名", "晶圆品名": "晶圆品名"},
    "赛卓-成品在制": {"规格": "产品规格", "品名": "产品品名", "晶圆品名": "晶圆型号"},
    "赛卓-成品库存": {"规格": "规格", "品名": "品名", "晶圆品名": "WAFER品名"},
    "赛卓-安全库存": {"规格": "OrderInformation", "品名": "ProductionNO.", "晶圆品名": "WaferID"},
    "赛卓-预测": {"规格": "产品型号", "品名": "ProductionNO.", "晶圆品名": "晶圆品名"}
}


class PivotProcessor:
    def process(self, uploaded_files: dict, output_buffer, additional_sheets: dict = None):
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            for filename, file_obj in uploaded_files.items():
                try:
                    df = pd.read_excel(file_obj)
                    config = CONFIG["pivot_config"].get(filename)
                    if not config:
                        st.warning(f"⚠️ 跳过未配置的文件：{filename}")
                        continue
    
                    sheet_name = filename[:30].replace(".xlsx", "")
                    st.write(f"📄 正在处理文件: `{filename}` → Sheet: `{sheet_name}`")
    
                    st.write(f"原始数据维度: {df.shape}")
                    st.dataframe(df.head(3))
    
                    # 日期处理
                    if "date_format" in config:
                        date_col = config["columns"]
                        df = self._process_date_column(df, date_col, config["date_format"])
    
                    # 映射替换（如果有）
                    if sheet_name in FIELD_MAPPINGS and "赛卓-新旧料号" in (additional_sheets or {}):
                        mapping_df = additional_sheets["赛卓-新旧料号"]
    
                        try:
                            mapping_df.columns = [
                                "旧规格", "旧品名", "旧晶圆品名",
                                "新规格", "新品名", "新晶圆品名",
                                "封装厂", "PC", "半成品"
                            ] + list(mapping_df.columns[9:])
                            st.success(f"✅ `{sheet_name}` 正在进行新旧料号替换...")
                        except Exception as e:
                            st.error(f"❌ `{sheet_name}` 替换前列名失败：{e}")
                            st.write("列名：", mapping_df.columns.tolist())
                            continue
    
                        df = apply_mapping_and_merge(df, mapping_df, FIELD_MAPPINGS[sheet_name])
    
                    # 构建透视表
                    pivoted = self._create_pivot(df, config)
                    st.write(f"✅ Pivot 表创建成功，维度：{pivoted.shape}")
                    st.dataframe(pivoted.head(3))
    
                    pivoted.to_excel(writer, sheet_name=sheet_name, index=False)
                    adjust_column_width(writer, sheet_name, pivoted)
    
                except Exception as e:
                    st.error(f"❌ 文件 `{filename}` 处理失败: {e}")
    
            # 写入附加 sheet（如预测、安全库存）
            if additional_sheets:
                for sheet_name, df in additional_sheets.items():
                    if sheet_name == "赛卓-新旧料号":
                        continue
                    try:
                        st.write(f"📎 正在写入附加表：{sheet_name}，数据维度：{df.shape}")
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        adjust_column_width(writer, sheet_name, df)
                    except Exception as e:
                        st.error(f"❌ 写入附加 Sheet `{sheet_name}` 失败: {e}")
    
        output_buffer.seek(0)

    def _process_date_column(self, df, date_col, date_format):
        if pd.api.types.is_numeric_dtype(df[date_col]):
            df[date_col] = df[date_col].apply(self._excel_serial_to_date)
        else:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

        new_col = f"{date_col}_年月"
        df[new_col] = df[date_col].dt.strftime(date_format)
        df[new_col] = df[new_col].fillna("未知日期")
        return df

    def _excel_serial_to_date(self, serial):
        try:
            return datetime(1899, 12, 30) + timedelta(days=float(serial))
        except:
            return pd.NaT

    def _create_pivot(self, df, config):
        config = config.copy()
        if "date_format" in config:
            config["columns"] = f"{config['columns']}_年月"

        pivoted = pd.pivot_table(
            df,
            index=config["index"],
            columns=config["columns"],
            values=config["values"],
            aggfunc=config["aggfunc"],
            fill_value=0
        )
        pivoted.columns = [f"{col[0]}_{col[1]}" if isinstance(col, tuple) else col for col in pivoted.columns]
        return pivoted.reset_index()
