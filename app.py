import io
import os
import openpyxl
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
不要な2ページ目以降（空ページやはみ出しページ）を排除し、必要なデータ行のみで1ページに収めて出力します。
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
            wb = openpyxl.load_workbook(base_file)

            for key, keywords in TARGET_SHEET_KEYWORDS:
                matched_sheet = find_matching_sheet(wb.sheetnames, keywords)
                if matched_sheet:
                    ws = wb[matched_sheet]
                    data_df_list = combined_rows[key]

                    if data_df_list:
                        all_data_df = pd.concat(
                            data_df_list, ignore_index=True
                        )
                        num_records = len(all_data_df)

                        if "No." in all_data_df.columns:
                            all_data_df["No."] = range(1, num_records + 1)

                        # 元の「合計行」の書式・数式・位置を取得
                        original_total_row = None
                        for r in range(9, ws.max_row + 1):
                            val = str(ws.cell(row=r, column=2).value).strip()
                            if val in ["合計", "計"]:
                                original_total_row = r
                                break

                        start_row = 9
                        # 1. 9行目から実データを書き込み
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

                        # 2. 合計行をデータ直下に詰めて移動＆SUM数式を補正
                        if original_total_row and original_total_row > new_total_row:
                            # 合計行のコピー
                            for c in range(1, ws.max_column + 1):
                                src_cell = ws.cell(row=original_total_row, column=c)
                                dst_cell = ws.cell(row=new_total_row, column=c)
                                if src_cell.value:
                                    val_str = str(src_cell.value)
                                    if val_str.startswith("=SUM("):
                                        col_let = get_column_letter(c)
                                        dst_cell.value = f"=SUM({col_let}9:{col_let}{last_data_row})"
                                    else:
                                        dst_cell.value = src_cell.value
                                if src_cell.has_style:
                                    dst_cell.font = src_cell.font.copy()
                                    dst_cell.border = src_cell.border.copy()
                                    dst_cell.fill = src_cell.fill.copy()
                                    dst_cell.alignment = src_cell.alignment.copy()
                                    dst_cell.number_format = src_cell.number_format

                            # 3. 合計行より下の「不要な未入力行（元々あった行）」を物理削除
                            delete_start = new_total_row + 1
                            delete_count = ws.max_row - delete_start + 1
                            if delete_count > 0:
                                ws.delete_rows(delete_start, amount=delete_count)

                            end_print_row = new_total_row
                        else:
                            end_print_row = last_data_row
                    else:
                        end_print_row = ws.max_row

                    # --- 🖨 1ページ内にピタッと収める印刷設定 ---
                    max_col_letter = get_column_letter(ws.max_column)
                    
                    # 印刷エリアを「A1〜最終データ列＆合計行」のみに限定
                    ws.print_area = f"A1:{max_col_letter}{end_print_row}"

                    # 1〜8行目をタイトル固定
                    ws.print_title_rows = "$1:$8"

                    # 用紙：A4・横向き
                    ws.page_setup.orientation = "landscape"
                    ws.page_setup.paperSize = ws.PAPERSIZE_A4

                    # 印刷倍率を「52%」に固定（右端の『投資効果予測』列まで1ページ幅にジャスト収める）
                    ws.page_setup.scale = 52

                    if ws.sheet_properties.pageSetUpPr:
                        ws.sheet_properties.pageSetUpPr.fitToPage = False

            output = io.BytesIO()
            wb.save(output)

            st.success(
                "🎉 不要ページを削除し、1ページ内に最適化された集約ファイルを作成しました！"
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
