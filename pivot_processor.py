import os
import pandas as pd
from datetime import datetime, timedelta
from config import CONFIG
from excel_utils import adjust_column_width

class PivotProcessor:
    def process(self, uploaded_files: dict, output_buffer, additional_sheets: dict = None):
        """
        主处理函数：生成透视表并写入 Excel
        - uploaded_files: 用户上传的主文件字典
        - output_buffer: BytesIO 输出缓冲区
        - additional_sheets: 额外需要写入的 DataFrame 表（如预测、安全库存、新旧料号）
        """
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            for filename, file_obj in uploaded_files.items():
                try:
                    df = pd.read_excel(file_obj)
                    config = CONFIG["pivot_config"].get(filename)
                    if not config:
                        continue

                    if "date_format" in config:
                        date_col = config["columns"]
                        df = self._process_date_column(df, date_col, config["date_format"])

                    pivoted = self._create_pivot(df, config)
                    sheet_name = filename[:30].replace(".xlsx", "")
                    pivoted.to_excel(writer, sheet_name=sheet_name, index=False)
                    adjust_column_width(writer, sheet_name, pivoted)
                except Exception as e:
                    print(f"{filename} 处理失败: {e}")

            # 写入附加 Sheet（预测、安全库存、新旧料号）
            if additional_sheets:
                for sheet_name, df in additional_sheets.items():
                    try:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        adjust_column_width(writer, sheet_name, df)
                        print(f"✅ 已写入附加 Sheet：{sheet_name}")
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
