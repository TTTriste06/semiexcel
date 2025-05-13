import os
import pandas as pd
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
                    continue

                sheet_name = filename[:30].replace(".xlsx", "")

                # 日期预处理（如果需要）
                if "date_format" in config:
                    date_col = config["columns"]
                    df = self._process_date_column(df, date_col, config["date_format"])

                # ⚠️ 如果在FIELD_MAPPINGS中，就执行替换逻辑
                if sheet_name in FIELD_MAPPINGS and "赛卓-新旧料号" in (additional_sheets or {}):
                    st.write(sheet_name)
                    mapping_df = additional_sheets["赛卓-新旧料号"]
                    df = apply_mapping_and_merge(df, mapping_df, FIELD_MAPPINGS[sheet_name])

                pivoted = self._create_pivot(df, config)
                pivoted.to_excel(writer, sheet_name=sheet_name, index=False)
                adjust_column_width(writer, sheet_name, pivoted)

            except Exception as e:
                print(f"{filename} 处理失败: {e}")

        # 附加 sheet（不透视）
        if additional_sheets:
            for sheet_name, df in additional_sheets.items():
                if sheet_name == "赛卓-新旧料号":
                    continue  # 不重复写入 mapping 表
                try:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    adjust_column_width(writer, sheet_name, df)
                except Exception as e:
                    print(f"❌ 写入 {sheet_name} 失败: {e}")

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
