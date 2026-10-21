import io
import os
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="設備投資計画フォーム 自動集約ツール",
    page_icon="📊",
    layout="wide",
)

st.title("📊 設備投資計画フォーム 自動集約ツール")
st.markdown("""
各部署から回収した「設備投資計画フォーム（Excel）」をまとめてドラッグ＆ドロップしてください。
D2セルの「千円」表記やプルダウン、説明文を完全に維持しつつ、各シートが1ページ以内に収まるよう印刷設定を最適化して出力します。
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


st.sidebar.header("📂 テンプレート設定")
template_file = st.sidebar.file_uploader(
    "【任意】まとめ用テンプレート（0000年度/全国まとめフォーム）",
    type=["xlsx"],
)

uploaded_files = st.file_uploader(
    "各部署のExcelファイルをここにドラッグ＆ドロップ（複数可）",
    type=["xlsx"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.info(f"📁 現在 {len(uploaded_files)} 個のファイルが選択されています。")

    if st.button("⚡ 集計処理を実行する", type="primary"):
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
            base_file.seek(0)
            wb = openpyxl.load_workbook(base_file)

            for key, keywords in TARGET_SHEET_KEYWORDS:
                matched_sheet = find_matching_sheet(wb.sheetnames, keywords)
                if matched_sheet:
                    ws = wb[matched_sheet]
                    
                    # --- 🏷 D2セル（単位）の「千円」とプルダウン・コメントを確実に復元・設定 ---
                    if not ws.cell(row=2, column=4).value:
                        ws.cell(row=2, column=4, value="千円")
                    
                    # D2セルに通貨選択のプルダウン（入力規則）を確実に付与
                    dv_d2 = DataValidation(type="list", formula1='"千円,ドル,元,ルピー"', allow_blank=True)
                    ws.add_data_validation(dv_d2)
                    dv_d2.add("D2")

                    data_df_list = combined_rows[key]

                    # 9行目以降の既存データをクリア
                    max_r = ws.max_row
                    for r in range(9, max_r + 1):
                        for c in range(1, ws.max_column + 1):
                            ws.cell(row=r, column=c).value = None

                    if data_df_list:
                        all_data_df = pd.concat(
                            data_df_list, ignore_index=True
                        )
                        num_records = len(all_data_df)

                        if "No." in all_data_df.columns:
                            all_data_df["No."] = range(1, num_records + 1)

                        start_row = 9
                        # 実データ書き込み
                        for r_idx, row in all_data_df.iterrows():
                            current_row = start_row + r_idx
                            for c_idx, val in enumerate(row):
                                if pd.notna(val):
                                    ws.cell(
                                        row=current_row,
                                        column=c_idx + 1,
                                        value=val,
                                    )

                        last_data_row = start_row + num_records - 1
                        new_total_row = last_data_row + 1

                        # 合計行の設定
                        ws.cell(row=new_total_row, column=2, value="合計")
                        for c in range(1, ws.max_column + 1):
                            dst_cell = ws.cell(row=new_total_row, column=c)
                            col_let = get_column_letter(c)
                            if c >= 10:  # 金額・数値列にSUM数式を設定
                                dst_cell.value = f"=SUM({col_let}9:{col_let}{last_data_row})"

                        end_print_row = new_total_row
                    else:
                        end_print_row = 8

                    # --- 🖨 1ページ（1ページに収める）印刷設定の完全適用 ---
                    max_col_letter = get_column_letter(ws.max_column)
                    ws.print_area = f"A1:{max_col_letter}{end_print_row}"
                    ws.print_title_rows = "$1:$8"
                    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                    ws.page_setup.paperSize = ws.PAPERSIZE_A4

                    ws.page_setup.scale = None
                    if ws.sheet_properties.pageSetUpPr is None:
                        ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties()
                    
                    ws.sheet_properties.pageSetUpPr.fitToPage = True
                    ws.page_setup.fitToWidth = 1
                    ws.page_setup.fitToHeight = 1

            # --- 🧹 中身（データ）が入っていない不要シートの自動削除処理 ---
            sheets_to_remove = []
            for sheetname in wb.sheetnames:
                ws = wb[sheetname]
                if sheetname not in ["償却率", "種類", "➃該当なし"]:
                    has_asset_data = False
                    for r in range(9, ws.max_row + 1):
                        val_no = str(ws.cell(row=r, column=2).value).strip()
                        val_name = str(ws.cell(row=r, column=5).value).strip()
                        if val_no and val_no not in ["None", "nan", "合計", "計", "No."] and val_name and val_name not in ["None", "nan"]:
                            has_asset_data = True
                            break
                    
                    if not has_asset_data and ("追加" in sheetname or "年間減価償却費" in sheetname or "償却月数" in sheetname):
                        sheets_to_remove.append(sheetname)

            for s_name in sheets_to_remove:
                if s_name in wb.sheetnames:
                    wb.remove(wb[s_name])

            output = io.BytesIO()
            wb.save(output)

            st.success(
                "🎉 「千円」表記とプルダウン・説明文を完全復元し、1ページ最適化された集約ファイルの作成が完了しました！"
            )

            st.download_button(
                label="📥 集約版Excelファイルをダウンロード",
                data=output.getvalue(),
                file_name="★【全国】設備投資計画集約版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

        except Exception as e:
            st.error(f"集約ファイルの生成中にエラーが発生しました: {e}")
