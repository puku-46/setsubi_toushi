import io
import os
import pandas as pd
import streamlit as st

# ページ基本設定
st.set_page_config(
    page_title="2027年度 設備投資計画フォーム 自動集約ツール",
    page_icon="📊",
    layout="wide",
)

st.title("📊 2027年度 設備投資計画フォーム 自動集約ツール")
st.markdown("""
各部署から回収した「2027年度設備投資計画フォーム（Excel）」をまとめてドラッグ＆ドロップしてください。
各シート（新規取得、補修・リース等）のデータを自動抽出・結合し、全国集約用の1つのExcelファイルを作成します。
""")

# 集約対象シート一覧
SHEETS_TO_PROCESS = [
    "①2027年度 新規取得・資本的支出用（年初）",
    "①2027年度 新規取得・資本的支出用（追加用）",
    "②2027年度 補修・リース（年初）",
    "②2027年度 補修・リース（追加用）",
    "⓷2026年度 新規取得・資本的支出用（追加・追加見込）",
]

# サイドバー設定
st.sidebar.header("⚙️ 設定オプション")
add_source_col = st.sidebar.checkbox(
    "集約データに「提出元ファイル名」列を追加する", value=True
)

# ファイルアップローダー
uploaded_files = st.file_uploader(
    "各部署のExcelファイルをここにドラッグ＆ドロップ（複数可）",
    type=["xlsx"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.info(f"📁 現在 {len(uploaded_files)} 個のファイルが選択されています。")

    if st.button("⚡ 集計処理を実行する", type="primary"):
        # 各シートごとのデータ格納用
        combined_data = {sheet: [] for sheet in SHEETS_TO_PROCESS}
        error_logs = []
        processed_count = 0

        progress_bar = st.progress(0)

        for idx, uploaded_file in enumerate(uploaded_files):
            filename = uploaded_file.name

            # 集約用まとめファイル自体が混ざっている場合はスキップ
            if "★" in filename or "全国" in filename or "集約" in filename:
                st.warning(
                    f"⚠️ スキップされました（集約用ファイルと検知）: {filename}"
                )
                continue

            try:
                # Excelファイルを読み込み
                xl = pd.ExcelFile(uploaded_file)

                for sheet in SHEETS_TO_PROCESS:
                    if sheet in xl.sheet_names:
                        # 8行目（header=7）をヘッダー行として取得
                        df = pd.read_excel(
                            uploaded_file, sheet_name=sheet, header=7
                        )

                        # 有効データの抽出基準（「資産名称」に値が入っている行）
                        if "資産名称" in df.columns:
                            df_valid = df.dropna(subset=["資産名称"]).copy()

                            # 不要な「合計」行などが混ざっている場合は除外
                            if "No." in df_valid.columns:
                                df_valid = df_valid[
                                    df_valid["No."].astype(str) != "合計"
                                ]

                            if not df_valid.empty:
                                if add_source_col:
                                    # 提出元ファイル名を列として挿入
                                    df_valid.insert(
                                        0,
                                        "提出元ファイル",
                                        os.path.splitext(filename)[0],
                                    )

                                combined_data[sheet].append(df_valid)

                processed_count += 1

            except Exception as e:
                error_logs.append(f"❌ {filename} の読み込みに失敗しました: {e}")

            # 進捗バー更新
            progress_bar.progress((idx + 1) / len(uploaded_files))

        # エラーログの表示
        if error_logs:
            st.error("以下のファイルで処理エラーが発生しました:")
            for err in error_logs:
                st.write(err)

        # 集約データの結合とExcelファイル化
        output = io.BytesIO()
        has_data = False

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet, dfs in combined_data.items():
                if dfs:
                    has_data = True
                    full_df = pd.concat(dfs, ignore_index=True)

                    # 通し番号（No.）の振り直し
                    if "No." in full_df.columns:
                        full_df["No."] = range(1, len(full_df) + 1)

                    # Excelのシートへ書き出し
                    full_df.to_excel(writer, sheet_name=sheet, index=False)

        if has_data:
            st.success(
                f"🎉 {processed_count} 個のファイルからデータの集約が完了しました！"
            )

            # タブで集約データのプレビュー表示
            st.subheader("🔍 集約データ プレビュー（先頭各5件）")
            preview_tabs = st.tabs([s[:10] for s in SHEETS_TO_PROCESS])

            for sheet, tab in zip(SHEETS_TO_PROCESS, preview_tabs):
                with tab:
                    if combined_data[sheet]:
                        preview_df = pd.concat(
                            combined_data[sheet], ignore_index=True
                        )
                        if "No." in preview_df.columns:
                            preview_df["No."] = range(1, len(preview_df) + 1)
                        st.dataframe(preview_df.head(5))
                    else:
                        st.caption("該当シートにデータはありませんでした。")

            # ダウンロードボタン表示
            st.download_button(
                label="📥 ★【全国】2027年度設備投資計画集約版.xlsx をダウンロード",
                data=output.getvalue(),
                file_name="★【全国】2027年度設備投資計画集約版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        else:
            st.warning(
                "抽出可能な有効データが見つかりませんでした。入力されたファイルをご確認ください。"
            )
