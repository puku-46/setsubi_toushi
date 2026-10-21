import io
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="設備投資計画フォーム 自動集約＆施設抽出ツール",
    page_icon="📊",
    layout="wide",
)

st.title("📊 設備投資計画フォーム 自動集約＆施設抽出ツール")
st.markdown("""
各部署から回収した「設備投資計画フォーム（Excel）」をまとめて集約します。  
集約後、**画面上のプルダウンから施設（部署）を複数選択**して、冷蔵部などの別ファイルとしてまとめて出力できます。
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

        # 集約結果をセッション状態に保存（画面再描画時にも保持するため）
        st.session_state["combined_rows"] = combined_rows
        st.session_state["processed_files_count"] = processed_count
        st.session_state["has_aggregated"] = True

# --- 集計完了後のUI表示 ---
if st.session_state.get("has_aggregated", False):
    combined_rows = st.session_state["combined_rows"]
    st.success("🎉 全部署のデータ集約が完了しました！")

    # 全データの中から C列（施設名・部署名）のユニークなリストを取得
    all_facilities = set()
    for key, df_list in combined_rows.items():
        for df in df_list:
            # 2番目の列（C列：部署、工場、冷蔵庫等）を取得
            dept_col = None
            for col in df.columns:
                if (
                    "部署" in str(col)
                    or "工場" in str(col)
                    or "冷蔵" in str(col)
                ):
                    dept_col = col
                    break
            if not dept_col and len(df.columns) >= 2:
                dept_col = df.columns[1]

            if dept_col and dept_col in df.columns:
                facs = (
                    df[dept_col]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .unique()
                    .tolist()
                )
                all_facilities.update([f for f in facs if f != "nan"])

    sorted_facilities = sorted(list(all_facilities))

    st.subheader("🏢 別ファイル（冷蔵部等）作成用の施設選択")
    st.markdown(
        "集約されたデータの中から、別ファイルに含めたい施設を選択してください。"
    )

    # 「冷蔵」が含まれるものをデフォルト選択肢として設定
    default_selected = [f for f in sorted_facilities if "冷蔵" in f]

    dept_file_label = st.text_input(
        "出力ファイル名（部署名）",
        value="冷蔵部",
        help="例：「冷蔵部」と入力すると『★【冷蔵部】設備投資計画集約版.xlsx』が出力されます。",
    )

    selected_facilities = st.multiselect(
        "抽出対象の施設・部署を選択してください（複数選択可）：",
        options=sorted_facilities,
        default=default_selected,
    )

    if st.button("📄 選択した施設でExcelファイルを生成する", type="secondary"):
        base_file = template_file if template_file else uploaded_files[0]

        try:
            # --- 1. 全国版のブック作成 ---
            base_file.seek(0)
            wb_national = openpyxl.load_workbook(base_file)

            # --- 2. 選択部署（冷蔵部）版のブック作成 ---
            base_file.seek(0)
            wb_dept = openpyxl.load_workbook(base_file)

            for key, keywords in TARGET_SHEET_KEYWORDS:
                matched_sheet = find_matching_sheet(
                    wb_national.sheetnames, keywords
                )
                if matched_sheet:
                    data_df_list = combined_rows[key]

                    # === 全国版の作成 ===
                    ws_nat = wb_national[matched_sheet]
                    if not ws_nat.cell(row=2, column=4).value:
                        ws_nat.cell(row=2, column=4, value="千円")
                    dv_nat = DataValidation(
                        type="list",
                        formula1='"千円,ドル,元,ルピー"',
                        allow_blank=True,
                    )
                    ws_nat.add_data_validation(dv_nat)
                    dv_nat.add("D2")

                    for r in range(9, ws_nat.max_row + 1):
                        for c in range(1, ws_nat.max_column + 1):
                            ws_nat.cell(row=r, column=c).value = None

                    if data_df_list:
                        all_data_df = pd.concat(
                            data_df_list, ignore_index=True
                        )
                        num_records = len(all_data_df)
                        if "No." in all_data_df.columns:
                            all_data_df["No."] = range(1, num_records + 1)

                        start_row = 9
                        for r_idx, row in all_data_df.iterrows():
                            current_row = start_row + r_idx
                            for c_idx, val in enumerate(row):
                                if pd.notna(val):
                                    ws_nat.cell(
                                        row=current_row,
                                        column=c_idx + 1,
                                        value=val,
                                    )

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

                    # 印刷設定（全国）
                    max_col_letter = get_column_letter(ws_nat.max_column)
                    ws_nat.print_area = (
                        f"A1:{max_col_letter}{end_print_row_nat}"
                    )
                    ws_nat.print_title_rows = "$1:$8"
                    ws_nat.page_setup.orientation = (
                        ws_nat.ORIENTATION_LANDSCAPE
                    )
                    ws_nat.page_setup.paperSize = ws_nat.PAPERSIZE_A4
                    ws_nat.page_setup.scale = None
                    if ws_nat.sheet_properties.pageSetUpPr is None:
                        ws_nat.sheet_properties.pageSetUpPr = (
                            openpyxl.worksheet.properties.PageSetupProperties()
                        )
                    ws_nat.sheet_properties.pageSetUpPr.fitToPage = True
                    ws_nat.page_setup.fitToWidth = 1
                    ws_nat.page_setup.fitToHeight = 1

                    # === 選択部署（冷蔵部）版の作成 ===
                    ws_dep = wb_dept[matched_sheet]
                    if not ws_dep.cell(row=2, column=4).value:
                        ws_dep.cell(row=2, column=4, value="千円")
                    dv_dep = DataValidation(
                        type="list",
                        formula1='"千円,ドル,元,ルピー"',
                        allow_blank=True,
                    )
                    ws_dep.add_data_validation(dv_dep)
                    dv_dep.add("D2")

                    for r in range(9, ws_dep.max_row + 1):
                        for c in range(1, ws_dep.max_column + 1):
                            ws_dep.cell(row=r, column=c).value = None

                    if data_df_list:
                        all_data_df = pd.concat(
                            data_df_list, ignore_index=True
                        )

                        dept_col_name = None
                        for col in all_data_df.columns:
                            if (
                                "部署" in str(col)
                                or "工場" in str(col)
                                or "冷蔵" in str(col)
                            ):
                                dept_col_name = col
                                break
                        if not dept_col_name and len(all_data_df.columns) >= 2:
                            dept_col_name = all_data_df.columns[1]

                        if dept_col_name and selected_facilities:
                            # 選択された施設リストに合致する行だけを抽出
                            dept_df = all_data_df[
                                all_data_df[dept_col_name]
                                .astype(str)
                                .str.strip()
                                .isin(selected_facilities)
                            ].copy()
                        else:
                            dept_df = pd.DataFrame(columns=all_data_df.columns)

                        num_records_dept = len(dept_df)
                        if num_records_dept > 0:
                            if "No." in dept_df.columns:
                                dept_df["No."] = range(
                                    1, num_records_dept + 1
                                )

                            start_row = 9
                            for r_idx, row in dept_df.reset_index(
                                drop=True
                            ).iterrows():
                                current_row = start_row + r_idx
                                for c_idx, val in enumerate(row):
                                    if pd.notna(val):
                                        ws_dep.cell(
                                            row=current_row,
                                            column=c_idx + 1,
                                            value=val,
                                        )

                            last_data_row = start_row + num_records_dept - 1
                            new_total_row = last_data_row + 1
                            ws_dep.cell(
                                row=new_total_row, column=2, value="合計"
                            )
                            for c in range(1, ws_dep.max_column + 1):
                                dst_cell = ws_dep.cell(
                                    row=new_total_row, column=c
                                )
                                col_let = get_column_letter(c)
                                if c >= 10:
                                    dst_cell.value = f"=SUM({col_let}9:{col_let}{last_data_row})"
                            end_print_row_dept = new_total_row
                        else:
                            end_print_row_dept = 8
                    else:
                        end_print_row_dept = 8

                    # 印刷設定（選択部署）
                    ws_dep.print_area = (
                        f"A1:{max_col_letter}{end_print_row_dept}"
                    )
                    ws_dep.print_title_rows = "$1:$8"
                    ws_dep.page_setup.orientation = (
                        ws_dep.ORIENTATION_LANDSCAPE
                    )
                    ws_dep.page_setup.paperSize = ws_dep.PAPERSIZE_A4
                    ws_dep.page_setup.scale = None
                    if ws_dep.sheet_properties.pageSetUpPr is None:
                        ws_dep.sheet_properties.pageSetUpPr = (
                            openpyxl.worksheet.properties.PageSetupProperties()
                        )
                    ws_dep.sheet_properties.pageSetUpPr.fitToPage = True
                    ws_dep.page_setup.fitToWidth = 1
                    ws_dep.page_setup.fitToHeight = 1

            # --- 不要シート削除（全国版） ---
            sheets_to_remove_nat = [
                s
                for s in wb_national.sheetnames
                if s not in ["償却率", "種類", "➃該当なし"]
                and not any(
                    str(wb_national[s].cell(row=r, column=2).value).strip()
                    not in ["None", "nan", "合計", "計", "No."]
                    for r in range(9, wb_national[s].max_row + 1)
                )
                and ("追加" in s or "年間減価償却費" in s or "償却月数" in s)
            ]
            for s_name in sheets_to_remove_nat:
                wb_national.remove(wb_national[s_name])

            # --- 不要シート削除（選択部署版） ---
            sheets_to_remove_dep = [
                s
                for s in wb_dept.sheetnames
                if s not in ["償却率", "種類", "➃該当なし"]
                and not any(
                    str(wb_dept[s].cell(row=r, column=2).value).strip()
                    not in ["None", "nan", "合計", "計", "No."]
                    for r in range(9, wb_dept[s].max_row + 1)
                )
                and ("追加" in s or "年間減価償却費" in s or "償却月数" in s)
            ]
            for s_name in sheets_to_remove_dep:
                wb_dept.remove(wb_dept[s_name])

            # ダウンロードバイナリの作成
            output_nat = io.BytesIO()
            wb_national.save(output_nat)

            output_dep = io.BytesIO()
            wb_dept.save(output_dep)

            st.divider()
            st.success("🎉 ファイルの生成が完了しました！")

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
                    label=f"📥 【{dept_file_label}版】抽出ファイルをダウンロード",
                    data=output_dep.getvalue(),
                    file_name=f"★【{dept_file_label}】設備投資計画集約版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

        except Exception as e:
            st.error(f"ファイル生成中にエラーが発生しました: {e}")
