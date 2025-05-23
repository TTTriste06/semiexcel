import os
import io
import re
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl import load_workbook
from config import CONFIG
from excel_utils import (
    adjust_column_width,
    clean_df,
    merge_header_for_summary, 
    mark_unmatched_keys_on_sheet,
    mark_keys_on_sheet,
    merge_duplicate_product_names,
    merge_duplicate_rows_by_key,
    clean_key_fields
)
from mapping_utils import apply_mapping_and_merge, apply_extended_substitute_mapping
from month_selector import process_history_columns
from summary import (
    merge_safety_inventory,
    append_unfulfilled_summary_columns,
    append_forecast_to_summary,
    merge_finished_inventory,
    append_product_in_progress
)

FIELD_MAPPINGS = {
    "赛卓-未交订单": {"规格": "规格", "品名": "品名", "晶圆品名": "晶圆品名"},
    "赛卓-成品在制": {"规格": "产品规格", "品名": "产品品名", "晶圆品名": "晶圆型号"},
    "赛卓-成品库存": {"规格": "规格", "品名": "品名", "晶圆品名": "WAFER品名"},
    "赛卓-安全库存": {"规格": "OrderInformation", "品名": "ProductionNO.", "晶圆品名": "WaferID"},
    "赛卓-预测": {"规格": "产品型号", "品名": "ProductionNO.", "晶圆品名": "晶圆品名"}
}

class PivotProcessor:
    def process(self, uploaded_files: dict, output_buffer, additional_sheets: dict = None):
        df_finished = pd.DataFrame()
        product_in_progress = pd.DataFrame()
        df_unfulfilled = pd.DataFrame()

        unmatched_safety = []
        unmatched_unfulfilled = []
        unmatched_forecast = []
        unmatched_finished = []
        unmatched_in_progress = []

        key_unfulfilled = []
        key_finished = []
        key_in_progress = []

        mapping_df = additional_sheets.get("赛卓-新旧料号", pd.DataFrame())
        
        all_mapped_keys = set()

        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            for filename, file_obj in uploaded_files.items():
                try:
                    df = pd.read_excel(file_obj)
                    df = clean_df(df)
                    config = CONFIG["pivot_config"].get(filename)
                    if not config:
                        st.warning(f"⚠️ 跳过未配置的文件：{filename}")
                        continue

                    sheet_name = filename.replace(".xlsx", "")

                    if sheet_name in FIELD_MAPPINGS and not mapping_df.empty:
                        mapping_df.columns = [
                            "旧规格", "旧品名", "旧晶圆品名",
                            "新规格", "新品名", "新晶圆品名",
                            "封装厂", "PC", "半成品", "备注",
                            "替代规格1", "替代品名1", "替代晶圆1", 
                            "替代规格2", "替代品名2", "替代晶圆2", 
                            "替代规格3", "替代品名3", "替代晶圆3",
                            "替代规格4", "替代品名5", "替代晶圆6"
                        ] + list(mapping_df.columns[22:])
                        st.success(f"✅ `{sheet_name}` 正在进行新旧料号替换...")

                        st.write(mapping_df)
                        df, mapped_keys = apply_mapping_and_merge(df, mapping_df, FIELD_MAPPINGS[sheet_name])
                        df, keys_sub = apply_extended_substitute_mapping(df, mapping_df, FIELD_MAPPINGS[sheet_name], None)
                        df = clean_key_fields(df, FIELD_MAPPINGS[sheet_name])
                        df = merge_duplicate_rows_by_key(df, FIELD_MAPPINGS[sheet_name])
                        all_mapped_keys.update(mapped_keys)

                        if sheet_name == "赛卓-未交订单":
                            key_unfulfilled = mapped_keys
                        elif sheet_name == "赛卓-成品库存":
                            key_finished = mapped_keys
                        elif sheet_name == "赛卓-成品在制":
                            key_in_progress = mapped_keys

                    if "date_format" in config:
                        df = self._process_date_column(df, config["columns"], config["date_format"])

                    pivoted = self._create_pivot(df, config)
                    pivoted.to_excel(writer, sheet_name=sheet_name, index=False)
                    adjust_column_width(writer, sheet_name, pivoted)

                    if sheet_name == "赛卓-未交订单":
                        df_unfulfilled = df
                        pivot_unfulfilled = pivoted
                    elif sheet_name == "赛卓-成品库存":
                        df_finished = pivoted
                    elif sheet_name == "赛卓-成品在制":
                        product_in_progress = pivoted

                except Exception as e:
                    st.error(f"❌ 文件 `{filename}` 处理失败: {e}")

            if df_unfulfilled.empty:
                st.error("❌ 缺少未交订单数据，无法构建汇总")
                return

            summary_preview = df_unfulfilled[["晶圆品名", "规格", "品名"]].drop_duplicates().reset_index(drop=True)

            try:
                if "赛卓-安全库存" in additional_sheets:
                    df_safety = additional_sheets["赛卓-安全库存"]
                    df_safety, keys_main = apply_mapping_and_merge(df_safety, mapping_df, FIELD_MAPPINGS["赛卓-安全库存"])
                    df_safety, keys_sub = apply_extended_substitute_mapping(df_safety, mapping_df, FIELD_MAPPINGS["赛卓-安全库存"], keys_main)
                    df_safety = merge_duplicate_rows_by_key(df_safety, FIELD_MAPPINGS["赛卓-安全库存"])
                    # all_mapped_keys.update(keys_main)
                    # all_mapped_keys.update(keys_sub)
                    summary_preview, unmatched_safety = merge_safety_inventory(summary_preview, df_safety)
                    st.success("✅ 已合并安全库存")
                    
                summary_preview, unmatched_unfulfilled = append_unfulfilled_summary_columns(summary_preview, pivot_unfulfilled)
                st.success("✅ 已合并未交订单")

                if "赛卓-预测" in additional_sheets:
                    forecast_df = additional_sheets["赛卓-预测"]
                    forecast_df.columns = forecast_df.iloc[0]
                    forecast_df = forecast_df[1:].reset_index(drop=True)
                    forecast_df, keys_main = apply_mapping_and_merge(forecast_df, mapping_df, FIELD_MAPPINGS["赛卓-预测"])
                    forecast_df, keys_sub = apply_extended_substitute_mapping(forecast_df, mapping_df, FIELD_MAPPINGS["赛卓-预测"], keys_main)
                    forecast_df = merge_duplicate_rows_by_key(forecast_df, field_map)
                    # all_mapped_keys.update(keys_main)
                    # all_mapped_keys.update(keys_sub)
                    summary_preview, unmatched_forecast = append_forecast_to_summary(summary_preview, forecast_df)
                    st.success("✅ 已合并预测数据")

                if not df_finished.empty:
                    summary_preview, unmatched_finished = merge_finished_inventory(summary_preview, df_finished)
                    st.success("✅ 已合并成品库存")

                if not product_in_progress.empty:
                    summary_preview, unmatched_in_progress = append_product_in_progress(summary_preview, product_in_progress, mapping_df)
                    st.success("✅ 已合并成品在制")

            except Exception as e:
                st.error(f"❌ 汇总数据合并失败: {e}")
                return

            summary_preview = summary_preview.drop_duplicates(subset=["晶圆品名", "规格", "品名"]).reset_index(drop=True)
            summary_preview = merge_duplicate_product_names(summary_preview)
            summary_preview.to_excel(writer, sheet_name="汇总", index=False)
            adjust_column_width(writer, "汇总", summary_preview)
            ws = writer.sheets["汇总"]

            header_row = list(summary_preview.columns)
            unfulfilled_cols = [col for col in header_row if "未交订单数量" in col or col in ("总未交订单", "历史未交订单数量")]
            forecast_cols = [col for col in header_row if "预测" in col]
            finished_cols = [col for col in header_row if col in ("数量_HOLD仓", "数量_成品仓", "数量_半成品仓")]

            merge_header_for_summary(
                ws, summary_preview,
                {
                    "安全库存": (" InvWaf", " InvPart"),
                    "未交订单": (unfulfilled_cols[0], unfulfilled_cols[-1]),
                    "预测": (forecast_cols[0], forecast_cols[-1]) if forecast_cols else ("", ""),
                    "成品库存": (finished_cols[0], finished_cols[-1]) if finished_cols else ("", ""),
                    "成品在制": ("成品在制", "半成品在制")
                }
            )

            for key, df in additional_sheets.items():
                df.to_excel(writer, sheet_name=key, index=False)
                adjust_column_width(writer, key, df)

            try:
                mark_unmatched_keys_on_sheet(writer.sheets["赛卓-安全库存"], unmatched_safety, wafer_col=1, spec_col=3, name_col=5)
                mark_unmatched_keys_on_sheet(writer.sheets["赛卓-未交订单"], unmatched_unfulfilled, wafer_col=1, spec_col=2, name_col=3)
                mark_unmatched_keys_on_sheet(writer.sheets["赛卓-预测"], unmatched_forecast, wafer_col=3, spec_col=1, name_col=2)
                writer.sheets["赛卓-预测"].delete_rows(2)
                mark_unmatched_keys_on_sheet(writer.sheets["赛卓-成品库存"], unmatched_finished, wafer_col=1, spec_col=2, name_col=3)
                mark_unmatched_keys_on_sheet(writer.sheets["赛卓-成品在制"], unmatched_in_progress, wafer_col=3, spec_col=4, name_col=5)
                writer.sheets["赛卓-新旧料号"].delete_rows(2)

                # mark_keys_on_sheet(writer.sheets["汇总"], all_mapped_keys, (2, 3, 1))
                # mark_keys_on_sheet(writer.sheets["赛卓-安全库存"], all_mapped_keys, (3, 5, 1))
                # mark_keys_on_sheet(writer.sheets["赛卓-未交订单"], all_mapped_keys, (2, 3, 1))
                # mark_keys_on_sheet(writer.sheets["赛卓-预测"], all_mapped_keys, (1, 2, 3))
                # mark_keys_on_sheet(writer.sheets["赛卓-成品库存"], all_mapped_keys, (2, 3, 1))
                # mark_keys_on_sheet(writer.sheets["赛卓-成品在制"], all_mapped_keys, (4, 5, 3))

                st.success("✅ 已完成未匹配项标记")
            except Exception as e:
                st.warning(f"⚠️ 未匹配标记失败：{e}")

            for name, ws in writer.sheets.items():
                col_letter = get_column_letter(ws.max_column)
                if name == "汇总":
                    ws.auto_filter.ref = f"A2:{col_letter}2"
                else:
                    ws.auto_filter.ref = f"A1:{col_letter}1"

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

        pivoted.columns = [f"{col[0]}_{col[1]}" if isinstance(col, tuple) else str(col) for col in pivoted.columns]

        if pd.Series(pivoted.columns).duplicated().any():
            from pandas.io.parsers import ParserBase
            original_cols = pivoted.columns
            deduped_cols = ParserBase({'names': original_cols})._maybe_dedup_names(original_cols)
            pivoted.columns = deduped_cols

        pivoted = pivoted.reset_index()

        if CONFIG.get("selected_month") and config.get("values") and "未交订单数量" in config.get("values"):
            st.info(f"📅 合并历史数据至：{CONFIG['selected_month']}")
            pivoted = process_history_columns(pivoted, config, CONFIG["selected_month"])
        return pivoted
