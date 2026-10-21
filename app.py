import io
import os
import pandas as pd
import openpyxl
import streamlit as st

# 自作のLINE通知モジュールをインポート
# (同じフォルダに line_notifier.py が配置されている前提です)
try:
    import line_notifier
except ImportError:
    line_notifier = None

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

# --- ファイルアップロード部分 ---
uploaded_files = st.file_uploader(
    "各部署のExcelファイルをここにドラッグ＆ドロップ（複数可）",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# テンプレート設定（任意）
with st.sidebar:
    st.header("テンプレート設定")
    template_file = st.file_uploader(
        "【任意】まとめ用テンプレート", type=["xlsx", "xls"]
    )

if uploaded_files:
    st.info(f"現在 {len(uploaded_files)} 個のファイルが選択されています。")
    
    # 全体集計ボタン
    if st.button("① まず全体集計を実行する", type="primary"):
        with st.spinner("集約処理を実行中です..."):
            try:
                # 実際の集約処理（ここに実際の結合ロジックを統合します）
                # ※ここではダミーまたは一般的な結合処理の枠組みを記載しています
                
                # サンプルとしてopenpyxlを使ったフッター（ページ番号）クリア処理の例
                # アップロードされた最初のファイルをベースに処理・調整するロジック
                output_buffer = io.BytesIO()
                
                # 例として最初のファイルを読み込んで処理
                wb = openpyxl.load_workbook(uploaded_files[0])
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    # ページ番号（フッター）を完全にクリアする処理
                    ws.oddFooter.left.text = ""
                    ws.oddFooter.center.text = ""
                    ws.oddFooter.right.text = ""
                    ws.evenFooter.left.text = ""
                    ws.evenFooter.center.text = ""
                    ws.evenFooter.right.text = ""
                
                wb.save(output_buffer)
                output_buffer.seek(0)
                
                # セッションステートに保存
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
            
            # --- LINE通知セクション ---
            st.markdown("---")
            st.subheader("💬 LINE通知テスト")
            notice_msg = st.text_input("送信メッセージ", value="【通知】設備投資計画の集約データが作成されました。")
            
            target_choice = st.selectbox(
                "通知先を選択", 
                options=["family", "user1"],
                format_func=lambda x: "家族グループ一斉同報" if x == "family" else "1さん個別通知"
            )
            
            if st.button("🚀 LINEに通知を送る"):
                if line_notifier:
                    success = line_notifier.send(notice_msg, target=target_choice)
                    if success:
                        st.success("LINEへの通知に成功しました！")
                    else:
                        st.error("LINEへの送信に失敗しました（ログまたはSecrets設定を確認してください）。")
                else:
                    st.warning("line_notifier モジュールが見つかりません。")

    with col2:
        st.subheader("🔍 ファイルB (個別抽出版)")
        st.markdown("転記したい名称（C列：部署名など）を選択してください：")
        # ダミーのセレクトボックス（必要に応じて実際の部署名リストに置き換えてください）
        selected_names = st.multiselect(
            "部署を選択",
            options=["宮城東洋", "通関課"],
            default=["宮城東洋", "通関課"]
        )

# --- データフレーム処理時の「No.の詰め直し」と「空行削除」の参考関数 ---
def clean_and_reindex_df(df):
    """
    資産名称などが空欄の行を除外して、No.を1から順番に振り直す関数
    """
    if "資産名称" in df.columns:
        # 資産名称が空欄（NaNや空白文字）の行を削除
        df = df.dropna(subset=["資産名称"])
        df = df[df["資産名称"].astype(str).str.strip() != ""]
    
    # No. 列が存在する場合は上から1, 2, 3... と振り直す
    if "No." in df.columns:
        df["No."] = range(1, len(df) + 1)
        
    return df
