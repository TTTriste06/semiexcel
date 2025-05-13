import os
import pandas as pd
from datetime import datetime, timedelta
from config import CONFIG
from excel_utils import adjust_column_width

class PivotProcessor:
    def __init__(self):
        self._validate_paths()

    def _validate_paths(self):
        if not os.path.exists(CONFIG["input_dir"]):
            raise FileNotFoundError(f"输入目录不存在: {CONFIG['input_dir']}")
        os.makedirs(os.path.dirname(CONFIG["output_file"]), exist_ok=True)

    def process(self):
        with pd.ExcelWriter(CONFIG["output_file"], engine='openpyxl') as writer:
            for file in os.listdir(CONFIG["input_dir"]):
                if not file.endswith('.xlsx') or file.startswith('~$'):
                    continue
                self._process_file(file, writer)
        print(f"处理完成，输出文件：{os.path.abspath(CONFIG['output_file'])}")

    def _process_file(self, filename, writer):
        try:
            filepath = os.path.join(CONFIG["input_dir"], filename)
            df = pd.read_excel(filepath, sheet_name=0)

            config = CONFIG["pivot_config"].get(filename)
            if not config:
                print(f"跳过未配置的文件: {filename}")
                return

            if "date_format" in config:
                date_col = config["columns"]
                df = self._process_date_column(df, date_col, config["date_format"])

            pivoted = self._create_pivot(df, config)

            sheet_name = filename[:30].rstrip('.xlsx')
            pivoted.to_excel(writer, sheet_name=sheet_name, index=False)
            adjust_column_width(writer, sheet_name, pivoted)
            print(f"成功处理: {filename}")

        except Exception as e:
            print(f"处理失败: {filename}，错误: {str(e)}")

    def _process_date_column(self, df, date_col, date_format):
        try:
            if pd.api.types.is_numeric_dtype(df[date_col]):
                print(f"检测到数值型日期，正在转换 Excel 序列号...")
                df[date_col] = df[date_col].apply(self._excel_serial_to_date)
            else:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')

            new_col = f"{date_col}_年月"
            df[new_col] = df[date_col].dt.strftime(date_format)

            invalid_mask = df[new_col].isnull()
            if invalid_mask.any():
                print(f"⚠️ 警告：发现 {invalid_mask.sum()} 条无效日期记录")
                df[new_col] = df[new_col].fillna("未知日期")

            valid_dates = df[date_col].dropna()
            if not valid_dates.empty:
                print(f"📅 日期范围: {valid_dates.min().date()} 至 {valid_dates.max().date()}")

            return df
        except KeyError:
            raise ValueError(f"日期列 [{date_col}] 在原始数据中不存在")

    def _excel_serial_to_date(self, serial):
        try:
            base_date = datetime(1899, 12, 30)
            return base_date + timedelta(days=float(serial))
        except:
            return pd.NaT

    def _create_pivot(self, df, config):
        if "date_format" in confi_
