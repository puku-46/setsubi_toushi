import io
import os
import pandas as pd
import openpyxl
import streamlit as st

# ページの基本設定
st.set_page_config(
    page_title="設備投資計画フォーム 自動集約ツール",
    page_icon="📊",
    layout="wide"
)

st.title("📊 設備投資計画フォーム 自動集約ツール")
st.markdown(
    "各部署から回収した「設備投資計画フォーム（Excel）」をまとめてドラッグ＆ドロップしてください。"
    "各シートが1ページ内に収まるよう印刷設定を最適化して出力します。"
)

# テンプレート設定（サイドバー）
with st.sidebar:
    st.header("テンプレート設定")
    template_file = st.file_uploader(
        "【任意】まとめ用テンプレート", type=["xlsx", "xls"]
    )

# --- ファイルアップロード部分 ---
uploaded_files = st.file_uploader(
    "各部署のExcelファイルをここにドラッグ＆ドロップ（複数可）",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"現在 {len(uploaded_files)} 個のファイルが選択されています。")
    
    # 全体集計を実行するボタン
    if st.button("① まず全体集計を実行する", type="primary"):
        with st.spinner("集約処理を実行中です..."):
            try:
                output_buffer = io.BytesIO()
                
                # アップロードされた最初のファイルを読み込んで処理
                # (※複数ファイルの結合や集約ロジックがすでにある場合は、そこに以下のクレンジング処理を適用します)
                wb = openpyxl.load_workbook(uploaded_files[0])
                
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    
                    # 1. 印刷設定：ページ番号（フッター）を完全にクリアする処理
                    ws.oddFooter.left.text = ""
                    ws.oddFooter.center.text = ""
                    ws.oddFooter.right.text = ""
                    ws.evenFooter.left.text = ""
                    ws.evenFooter.center.text = ""
                    ws.evenFooter.right.text = ""
                
                wb.save(output_buffer)
                output_buffer.seek(0)
                
                # 処理結果をセッションに保存
                st.session_state["processed_file"] = output_buffer.getvalue()
                st.session_state["集約完了"] = True
                
                st.success("🎉 全体の集約用データ（ファイルA）の作成が完了しました！")
                
            except Exception as e:
                st.error(f"集約処理中にエラーが発生しました: {e}")

# --- 集約完了後の表示エリア ---
if st.session_state.get("集約完了", False):
    st.divider()
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📦 ファイルA (全体版)")
        if "processed_file" in st.session_state:
            st.download_button(
                label="📥 全体の集約版をダウンロード",
                data=st.session_state["processed_file"],
                file_name="設備投資計画_集約済み.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

    with col2:
        st.subheader("🔍 ファイルB (個別抽出版)")
        st.markdown("転記したい名称（C列：部署名など）を選択してください：")
        selected_names = st.multiselect(
            "部署を選択",
            options=["宮城東洋", "通関課"],
            default=["宮城東洋", "通関課"]
        )

# --- pandas使用時のデータ整理用関数（必要に応じて集約ロジック内で呼び出してください） ---
def clean_and_reindex_df(df):
    """
    資産名称などが空欄の行を除外して、No.を1から順番に詰め直す関数
    """
    if "資産名称" in df.columns:
        # 資産名称が空欄（NaNや空白文字）の行を削除
        df = df.dropna(subset=["資産名称"])
        df = df[df["資産名称"].astype(str).str.strip() != ""]
    
    # No. 列が存在する場合は上から1, 2, 3... と綺麗に振り直す
    if "No." in df.columns:
        df["No."] = range(1, len(df) + 1)
        
    return df
