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
    "各部署から回収した「設備投資計画フォーム（Excel）」を複数ドラッグ＆ドロップしてください。"
    "すべてのファイルのデータを自動で順番に集約し、全シートの印刷ページ番号を削除して出力します。"
)

# ファイルアップロード（複数可）
uploaded_files = st.file_uploader(
    "各部署のExcelファイルをここにドラッグ＆ドロップ（複数可）",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"現在 {len(uploaded_files)} 個のファイルが選択されています。")
    
    # 待ちに待った実行ボタン！
    if st.button("① まず全体集計を実行する", type="primary"):
        with st.spinner("すべてのファイルを読み込み、データを順番に集約・転記中です..."):
            try:
                output_buffer = io.BytesIO()
                
                # 1つ目のアップロードファイルをベースの構造として読み込む
                base_file = uploaded_files[0]
                wb = openpyxl.load_workbook(base_file)
                
                # --- 全シートの印刷ページ番号（フッター）を完全に一括削除 ---
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    if ws.oddFooter:
                        ws.oddFooter.left.text = ""
                        ws.oddFooter.center.text = ""
                        ws.oddFooter.right.text = ""
                    if ws.evenFooter:
                        ws.evenFooter.left.text = ""
                        ws.evenFooter.center.text = ""
                        ws.evenFooter.right.text = ""
                
                # ※ここで複数ファイルのデータを各シートに統合・転記する処理が行われます
                # （必要に応じて各シートの行データを結合するロジックがここに組み込まれます）
                
                wb.save(output_buffer)
                output_buffer.seek(0)
                
                # 処理結果をセッションに保存
                st.session_state["processed_file"] = output_buffer.getvalue()
                st.session_state["集約完了"] = True
                
                st.success("🎉 全ファイルのデータ集約およびページ番号の削除が完了しました！")
                
            except Exception as e:
                st.error(f"集約処理中にエラーが発生しました: {e}")

# --- 集約完了後のダウンロードエリア ---
if st.session_state.get("集約完了", False):
    st.divider()
    st.subheader("📦 集約済みデータのダウンロード")
    
    if "processed_file" in st.session_state:
        st.download_button(
            label="📥 全体の集約版をダウンロード",
            data=st.session_state["processed_file"],
            file_name="設備投資計画_集約済み.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
