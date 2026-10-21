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
各シートが必ず1ページ以内に収まるよう印刷設定を最適化して出力します。
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

# --- 状態保持（Session State）の初期化 ---
if "combined_rows_A" not in st.session_state:
    st.session_state.combined_rows_A = None
if "c_column_unique_names" not in st.session_state:
    st.session_state.c_column_unique_names = []

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

# Excel作成共通処理（openpyxl）
def create_excel_file(source_combined_rows):
    base_file = template_file if template_file else uploaded_files
    if hasattr(base_file, "seek"):
        base_file.seek(0)
    wb = openpyxl.load_workbook(base_file)
    
    for key, keywords in TARGET_SHEET_KEYWORDS:
        matched_sheet = find_matching_sheet(wb.sheetnames, keywords)
        if matched_sheet:
            ws = wb[matched_sheet]
            data_df_list = source_combined_rows[key]
            
            if data_df_list:
                all_data_df = pd.concat(data_df_list, ignore_index=True)
                num_records = len(all_data_df)
                if "No." in all_data_df.columns:
                    all_data_df["No."] = range(1, num_records + 1)
                
                original_total_row = None
                for r in range(9, ws.max_row + 1):
                    val = str(ws.cell(row=r, column=2).value).strip()
                    if val in ["合計", "計"]:
                        original_total_row = r
                        break
                
                start_row = 9
                for r_idx, row in all_data_df.iterrows():
                    current_row = start_row + r_idx
                    for c_idx, val in enumerate(row):
                        if pd.notna(val):
                            ws.cell(row=current_row, column=c_idx + 1, value=val)
                
                last_data_row = start_row + num_records - 1
                new_total_row = last_data_row + 1
                
                if original_total_row and original_total_row > new_total_row:
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
                    
                    delete_start = new_total_row + 1
                    delete_count = ws.max_row - delete_start + 1
                    if delete_count > 0:
                        ws.delete_rows(delete_start, amount=delete_count)
                    end_print_row = new_total_row
                else:
                    end_print_row = last_data_row
            else:
                # データが空の場合、データ行(9行目以降)をクリア
                original_total_row = None
                for r in range(9, ws.max_row + 1):
                    val = str(ws.cell(row=r, column=2).value).strip()
                    if val in ["合計", "計"]:
                        original_total_row = r
                        break
                if original_total_row and original_total_row > 9:
                    for c in range(1, ws.max_column + 1):
                        ws.cell(row=9, column=c).value = ws.cell(row=original_total_row, column=c).value
                        src = ws.cell(row=original_total_row, column=c)
                        dst = ws.cell(row=9, column=c)
                        if src.has_style:
                            dst.font = src.font.copy(); dst.border = src.border.copy(); dst.fill = src.fill.copy(); dst.alignment = src.alignment.copy(); dst.number_format = src.number_format
                    ws.delete_rows(10, amount=ws.max_row - 9)
                    end_print_row = 9
                else:
                    end_print_row = ws.max_row
            
            # 印刷設定
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
            
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

if uploaded_files:
    st.info(f"📁 現在 {len(uploaded_files)} 個のファイルが選択されています。")
    
    if st.button("⚡ ① まず全体集計を実行する", type="primary"):
        combined_rows_A = {key: [] for key, _ in TARGET_SHEET_KEYWORDS}
        all_c_names = set()
        
        for uploaded_file in uploaded_files:
            filename = uploaded_file.name
            if any(kw in filename for kw in ["★", "全国", "集約", "0000年度"]):
                continue
            try:
                xl = pd.ExcelFile(uploaded_file)
                for key, keywords in TARGET_SHEET_KEYWORDS:
                    matched_sheet = find_matching_sheet(xl.sheet_names, keywords)
                    if matched_sheet:
                        df = pd.read_excel(uploaded_file, sheet_name=matched_sheet, header=7)
                        if "資産名称" in df.columns:
                            df_valid = df.dropna(subset=["資産名称"]).copy()
                            if "No." in df_valid.columns:
                                df_valid = df_valid[df_valid["No."].astype(str) != "合計"]
                            
                            if not df_valid.empty:
                                combined_rows_A[key].append(df_valid)
                                # C列（資産名称列）のユニークな値を収集
                                for val in df_valid["資産名称"].dropna().unique():
                                    all_c_names.add(str(val).strip())
            except Exception as e:
                st.error(f"❌ {filename} の読み込みエラー: {e}")
        
        st.session_state.combined_rows_A = combined_rows_A
        st.session_state.c_column_unique_names = sorted(list(all_c_names))

# 全体集計が完了している場合のみ、次のステップを表示
if st.session_state.combined_rows_A:
    st.success("🎉 全体の集約用データ（ファイルA）の作成が完了しました！")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📦 ファイルA (全体版)")
        file_A_data = create_excel_file(st.session_state.combined_rows_A)
        st.download_button(
            label="📥 全体の集約版をダウンロード",
            data=file_A_data,
            file_name="★【全国】設備投資計画集約版.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="download_all"
        )
        
    with col2:
        st.subheader("🔍 ファイルB (個別抽出版)")
        if st.session_state.c_column_unique_names:
            # 【変更点】 st.multiselect に変更し、複数選択可能に
            selected_names = st.multiselect(
                "転記したい名称（C列）をすべて選択してください（複数可）：",
                options=st.session_state.c_column_unique_names,
                placeholder="名称を選択または検索..."
            )
            
            if selected_names:
                # 選択された複数の名称のいずれかに一致（isin）するデータを抽出
                combined_rows_B = {key: [] for key, _ in TARGET_SHEET_KEYWORDS}
                for key, dfs in st.session_state.combined_rows_A.items():
                    for df in dfs:
                        # トリム（空白削除）して比較
                        df_filtered = df[df["資産名称"].astype(str).str.strip().isin(selected_names)].copy()
                        if not df_filtered.empty:
                            combined_rows_B[key].append(df_filtered)
                
                # ファイルBの作成とダウンロードボタン表示
                file_B_data = create_excel_file(combined_rows_B)
                st.download_button(
                    label=f"📥 選択した {len(selected_names)} 件を転記したファイルBをダウンロード",
                    data=file_B_data,
                    file_name="★【抽出】設備投資計画_複数選択版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_filtered"
                )
            else:
                st.info("💡 名称を1つ以上選択すると、ファイルBのダウンロードボタンが表示されます。")
        else:
            st.warning("C列（資産名称）にデータが見つかりませんでした。")
