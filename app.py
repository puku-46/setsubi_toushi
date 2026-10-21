import io
import os
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="設備投資計画フォーム 自動集約＆冷蔵部抽出ツール",
    page_icon="📊",
    layout="wide",
)

st.title("📊 設備投資計画フォーム 自動集約＆冷蔵部抽出ツール")
st.markdown("""
各部署から回収した「設備投資計画フォーム（Excel）」をまとめてドラッグ＆ドロップしてください。
全国の集約ファイルと、指定した部署（冷蔵部など）を抽出した別ファイルの**2種類を同時に作成**します。
""")

TARGET_SHEET_KEYWORDS = [
    ("新規取得_年初", ["①", "新規取得", "年初"]),
    ("新規取得_追加", ["①", "新規取得", "追加用"]),
    ("補修リース_年初", ["②", "補修", "年初"]),
    ("補修リース_追加", ["②", "補修", "追加用"]),
    ("前年度追加", ["⓷", "新規取得", "追加"]),
]


def find_matching_sheet(sheet_list, keywords):
    for s in sheet_list:
        if all(kw in s for kw in keywords):
            return s
    return None


st.sidebar.header("📂 設定メニュー")
template_file = st.sidebar.file_uploader(
    "【任意】まとめ用テンプレート（0000年度/全国まとめフォーム）",
    type=["xlsx"],
)

target_dept_keyword = st.sidebar.text_input(
    "抽出する部署キーワード（別ファイル用）",
    value="冷蔵部",
    help="C列（部署、工場、冷蔵庫等）に含まれる文字列で冷蔵部用ファイルを抽出します。",
)

uploaded_files = st.file_uploader(
    "各部署のExcelファイルをここにドラッグ＆ドロップ（複数可）",
    type=["xlsx"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.info(f"📁 現在 {len(uploaded_files)} 個のファイルが選択されています。")

    if st.button("⚡ 集計＆別ファイル作成を実行する", type="primary"):
        combined_rows = {key: [] for key, _ in TARGET_SHEET_KEYWORDS}
        processed_count = 0

        for uploaded_file in uploaded_files:
            filename = uploaded_file.name
            if (
                "★" in filename
                or "全国" in filename
                or "集約" in filename
                or "0000年度" in filename
            ):
                continue

            try:
                xl = pd.ExcelFile(uploaded_file)
                for key, keywords in TARGET_SHEET_KEYWORDS:
                    matched_sheet = find_matching_sheet(
                        xl.sheet_names, keywords
                    )
                    if matched_sheet:
                        df = pd.read_excel(
                            uploaded_file, sheet_name=matched_sheet, header=7
                        )
                        if "資産名称" in df.columns:
                            df_valid = df.dropna(subset=["資産名称"]).copy()
                            if "No." in df_valid.columns:
                                df_valid = df_valid[
                                    df_valid["No."].astype(str) != "合計"
                                ]

                            if not df_valid.empty:
                                combined_rows[key].append(df_valid)

                processed_count += 1
            except Exception as e:
                st.error(f"❌ {filename} の読み込みエラー: {e}")

        base_file = template_file if template_file else uploaded_files[0]

        try:
            # --- 1. 全国版のブック作成 ---
            base_file.seek(0)
            wb_national = openpyxl.load_workbook(base_file)

            # --- 2. 冷蔵部版のブック作成（テンプレートをもう一つ読み込む） ---
            base_file.seek(0)
            wb_dept = openpyxl.load_workbook(base_file)

            for key, keywords in TARGET_SHEET_KEYWORDS:
                matched_sheet = find_matching_sheet(wb_national.sheetnames, keywords)
                if matched_sheet:
                    # --- 全国版の処理 ---
                    ws_nat = wb_national[matched_sheet]
                    if not ws_nat.cell(row=2, column=4).value:
                        ws_nat.cell(row=2, column=4, value="千円")
                    dv_nat = DataValidation(type="list", formula1='"千円,ドル,元,ルピー"', allow_blank=True)
                    ws_nat.add_data_validation(dv_nat)
                    dv_nat.add("D2")

                    data_df_list = combined_rows[key]

                    # 9行目以降クリア（全国版）
                    for r in range(9, ws_nat.max_row + 1):
                        for c in range(1, ws_nat.max_column + 1):
                            ws_nat.cell(row=r, column=c).value = None

                    if data_df_list:
                        all_data_df = pd.concat(data_df_list, ignore_index=True)
                        num_records = len(all_data_df)
                        if "No." in all_data_df.columns:
                            all_data_df["No."] = range(1, num_records + 1)

                        start_row = 9
                        for r_idx, row in all_data_df.iterrows():
                            current_row = start_row + r_idx
                            for c_idx, val in enumerate(row):
                                if pd.notna(val):
                                    ws_nat.cell(row=current_row, column=c_idx + 1, value=val)

                        last_data_row = start_row + num_records - 1
                        new_total_row = last_data_row + 1
                        ws_nat.cell(row=new_total_row, column=2, value="合計")
                        for c in range(1, ws_nat.max_column + 1):
                            dst_cell = ws_nat.cell(row=new_total_row, column=c)
                            col_let = get_column_letter(c)
                            if c >= 10:
                                dst_cell.value = f"=SUM({col_let}9:{col_let}{last_data_row})"
                        end_print_row_nat = new_total_row
                    else:
                        end_print_row_nat = 8

                    # 印刷設定（全国版）
                    max_col_letter = get_column_letter(ws_nat.max_column)
                    ws_nat.print_area = f"A1:{max_col_letter}{end_print_row_nat}"
                    ws_nat.print_title_rows = "$1:$8"
                    ws_nat.page_setup.orientation = ws_nat.ORIENTATION_LANDSCAPE
                    ws_nat.page_setup.paperSize = ws_nat.PAPERSIZE_A4
                    ws_nat.page_setup.scale = None
                    if ws_nat.sheet_properties.pageSetUpPr is None:
                        ws_nat.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties()
                    ws_nat.sheet_properties.pageSetUpPr.fitToPage = True
                    ws_nat.page_setup.fitToWidth = 1
                    ws_nat.page_setup.fitToHeight = 1


                    # --- 冷蔵部版の処理（指定キーワードでフィルタリング） ---
                    ws_dep = wb_dept[matched_sheet]
                    if not ws_dep.cell(row=2, column=4).value:
                        ws_dep.cell(row=2, column=4, value="千円")
                    dv_dep = DataValidation(type="list", formula1='"千円,ドル,元,ルピー"', allow_blank=True)
                    ws_dep.add_data_validation(dv_dep)
                    dv_dep.add("D2")

                    # 9行目以降クリア（冷蔵部版）
                    for r in range(9, ws_dep.max_row + 1):
                        for c in range(1, ws_dep.max_column + 1):
                            ws_dep.cell(row=r, column=c).value = None

                    if data_df_list:
                        all_data_df = pd.concat(data_df_list, ignore_index=True)
                        
                        # C列（インデックス2、列番号3「部署、工場、冷蔵庫等」）に target_dept_keyword が含まれる行を抽出
                        dept_col_name = None
                        for col in all_data_df.columns:
                            if "部署" in str(col) or "工場" in str(col) or "冷蔵" in str(col):
                                dept_col_name = col
                                break
                        if not dept_col_name and len(all_data_df.columns) >= 3:
                            dept_col_name = all_data_df.columns[2] # 通常C列（0始まりで2）

                        if dept_col_name:
                            dept_df = all_data_df[all_data_df[dept_col_name].astype(str).str.contains(target_dept_keyword, na=False)].copy()
                        else:
                            dept_df = pd.DataFrame(columns=all_data_df.columns)

                        num_records_dept = len(dept_df)
                        if num_records_dept > 0:
                            if "No." in dept_df.columns:
                                dept_df["No."] = range(1, num_records_dept + 1)

                            start_row = 9
                            for r_idx, row in dept_df.reset_index(drop=True).iterrows():
                                current_row = start_row + r_idx
                                for c_idx, val in enumerate(row):
                                    if pd.notna(val):
                                        ws_dep.cell(row=current_row, column=c_idx + 1, value=val)

                            last_data_row = start_row + num_records_dept - 1
                            new_total_row = last_data_row + 1
                            ws_dep.cell(row=new_total_row, column=2, value="合計")
                            for c in range(1, ws_dep.max_column + 1):
                                dst_cell = ws_dep.cell(row=new_total_row, column=c)
                                col_let = get_column_letter(c)
                                if c >= 10:
                                    dst_cell.value = f"=SUM({col_let}9:{col_let}{last_data_row})"
                            end_print_row_dept = new_total_row
                        else:
                            end_print_row_dept = 8
                    else:
                        end_print_row_dept = 8

                    # 印刷設定（冷蔵部版）
                    ws_dep.print_area = f"A1:{max_col_letter}{end_print_row_dept}"
                    ws_dep.print_title_rows = "$1:$8"
                    ws_dep.page_setup.orientation = ws_dep.ORIENTATION_LANDSCAPE
                    ws_dep.page_setup.paperSize = ws_dep.PAPERSIZE_A4
                    ws_dep.page_setup.scale = None
                    if ws_dep.sheet_properties.pageSetUpPr is None:
                        ws_dep.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties()
                    ws_dep.sheet_properties.pageSetUpPr.fitToPage = True
                    ws_dep.page_setup.fitToWidth = 1
                    ws_dep.page_setup.fitToHeight = 1

            # --- 不要シートの削除（全国版） ---
            sheets_to_remove_nat = []
            for sheetname in wb_national.sheetnames:
                ws = wb_national[sheetname]
                if sheetname not in ["償却率", "種類", "➃該当なし"]:
                    has_asset_data = any(
                        str(ws.cell(row=r, column=2).value).strip() not in ["None", "nan", "合計", "計", "No."] and
                        str(ws.cell(row=r, column=5).value).strip() not in ["None", "nan"]
                        for r in range(9, ws.max_row + 1)
                    )
                    if not has_asset_data and ("追加" in sheetname or "年間減価償却費" in sheetname or "償却月数" in sheetname):
                        sheets_to_remove_nat.append(sheetname)
            for s_name in sheets_to_remove_nat:
                if s_name in wb_national.sheetnames:
                    wb_national.remove(wb_national[s_name])

            # --- 不要シートの削除（冷蔵部版） ---
            sheets_to_remove_dep = []
            for sheetname in wb_dept.sheetnames:
                ws = wb_dept[sheetname]
                if sheetname not in ["償却率", "種類", "➃該当なし"]:
                    has_asset_data = any(
                        str(ws.cell(row=r, column=2).value).strip() not in ["None", "nan", "合計", "計", "No."] and
                        str(ws.cell(row=r, column=5).value).strip() not in ["None", "nan"]
                        for r in range(9, ws.max_row + 1)
                    )
                    if not has_asset_data and ("追加" in sheetname or "年間減価償却費" in sheetname or "償却月数" in sheetname):
                        sheets_to_remove_dep.append(sheetname)
            for s_name in sheets_to_remove_dep:
                if s_name in wb_dept.sheetnames:
                    wb_dept.remove(wb_dept[s_name])

            # --- バイナリ出力の作成 ---
            output_nat = io.BytesIO()
            wb_national.save(output_nat)

            output_dep = io.BytesIO()
            wb_dept.save(output_dep)

            st.success("🎉 全国集計および「" + target_dept_keyword + "」抽出ファイルの作成が完了しました！")

            # 2つのダウンロードボタンを並べて表示
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 【全国版】集約ファイルをダウンロード",
                    data=output_nat.getvalue(),
                    file_name="★【全国】設備投資計画集約版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            with col2:
                st.download_button(
                    label=f"📥 【{target_dept_keyword}版】抽出ファイルをダウンロード",
                    data=output_dep.getvalue(),
                    file_name=f"★【{target_dept_keyword}】設備投資計画集約版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

        except Exception as e:
            st.error(f"ファイルの生成中にエラーが発生しました: {e}")
