import streamlit as st
from datetime import datetime
from models import PatrolData
from config import Config
from excel.writer import ExcelWriter

def main():
    st.set_page_config(
        page_title="日報作成ツール",
        page_icon="📋",
        layout="wide"
    )
    
    st.title("📋 日報作成ツール")
    st.markdown("---")
    
    if 'config' not in st.session_state:
        st.session_state.config = Config()
    
    config = st.session_state.config
    
    tab1, tab2 = st.tabs(["📝 日報作成", "👥 スタッフ管理"])
    
    with tab1:
        st.header("日報作成")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("担当者選択")
            
            post4_options = [""] + config.security_staff_list
            post4 = st.selectbox("4ポスト担当", post4_options, key="post4")
            
            post5_options = [""] + config.security_staff_list
            post5 = st.selectbox("5ポスト担当", post5_options, key="post5")
            
            post1_options = [""] + config.security_staff_list
            post1 = st.selectbox("1ポスト担当", post1_options, key="post1")
            
            supervisor_options = [""] + config.facility_staff_list
            supervisor = st.selectbox("責任者・巡回担当", supervisor_options, key="supervisor")
            
            st.subheader("巡回設定")
            patrol_start = st.selectbox(
                "巡回開始時刻", 
                ["21:00頃", "22:00頃"], 
                key="patrol_start"
            )
        
        with col2:
            st.subheader("劇場使用状況")
            large_theater = st.checkbox("大劇場使用", key="large_theater")
            medium_theater = st.checkbox("中劇場（楽屋）使用", key="medium_theater")
            small_theater = st.checkbox("小劇場使用", key="small_theater")
            
            st.subheader("Excelファイル")
            uploaded_file = st.file_uploader(
                "日報テンプレートファイルをアップロード",
                type=['xlsx'],
                help="日報のテンプレートExcelファイルを選択してください（最大10MB）"
            )
            
            # ファイル検証
            if uploaded_file is not None:
                # ファイルサイズチェック（10MB制限）
                if uploaded_file.size > 10 * 1024 * 1024:
                    st.error("ファイルサイズが大きすぎます。10MB以下のファイルを選択してください。")
                    st.stop()  # 処理を停止
                else:
                    st.success(f"ファイル '{uploaded_file.name}' が正常にアップロードされました")
                    st.info(f"ファイルサイズ: {uploaded_file.size / 1024:.1f} KB")
        
        st.markdown("---")
        
        if st.button("📋 日報作成", type="primary", use_container_width=True):
            if not all([post4, post5, post1, supervisor]):
                st.error("すべての担当者を選択してください。")
            elif not uploaded_file:
                st.error("Excelファイルをアップロードしてください。")
            else:
                # 担当者の重複チェック
                staff_list = [post4, post5, post1, supervisor]
                if len(set(staff_list)) != len(staff_list):
                    st.error("同じ担当者が複数のポストに割り当てられています。担当者を変更してください。")
                else:
                    try:
                        patrol_data = PatrolData(
                            post4=post4,
                            post5=post5,
                            post1=post1,
                            supervisor=supervisor,
                            patrol_start=patrol_start,
                            large_theater_used=large_theater,
                            medium_theater_used=medium_theater,
                            small_theater_used=small_theater
                        )
                        
                        writer = ExcelWriter()
                        file_bytes = uploaded_file.read()
                        
                        with st.spinner("日報を作成中..."):
                            output_bytes = writer.write_report(file_bytes, patrol_data)
                        
                        today = datetime.today()
                        filename = f"日報_{today.strftime('%Y%m%d')}.xlsx"
                        
                        st.success("日報が正常に作成されました！")
                        st.download_button(
                            label="📥 日報をダウンロード",
                            data=output_bytes,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
                        
                    except Exception as e:
                        st.error(f"エラーが発生しました: {e}")
    
    with tab2:
        st.header("スタッフ管理")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("警備担当者")
            
            new_security = st.text_input("新しい警備担当者を追加", key="new_security")
            if st.button("追加", key="add_security"):
                if new_security:
                    if config.add_security_staff(new_security):
                        st.success(f"'{new_security}' を追加しました。")
                        # 入力フィールドをクリア
                        st.session_state.new_security = ""
                    else:
                        st.warning("既に登録されているか、無効な名前です。")
                else:
                    st.warning("名前を入力してください。")
            
            st.markdown("**登録済み警備担当者:**")
            for staff in config.security_staff_list:
                col_name, col_delete = st.columns([3, 1])
                with col_name:
                    st.text(staff)
                with col_delete:
                    if st.button("削除", key=f"del_sec_{staff}"):
                        config.remove_security_staff(staff)
                        st.success(f"'{staff}' を削除しました。")
                        st.rerun()  # リスト更新のため必要
        
        with col2:
            st.subheader("設備担当者")
            
            new_facility = st.text_input("新しい設備担当者を追加", key="new_facility")
            if st.button("追加", key="add_facility"):
                if new_facility:
                    if config.add_facility_staff(new_facility):
                        st.success(f"'{new_facility}' を追加しました。")
                        # 入力フィールドをクリア
                        st.session_state.new_facility = ""
                    else:
                        st.warning("既に登録されているか、無効な名前です。")
                else:
                    st.warning("名前を入力してください。")
            
            st.markdown("**登録済み設備担当者:**")
            for staff in config.facility_staff_list:
                col_name, col_delete = st.columns([3, 1])
                with col_name:
                    st.text(staff)
                with col_delete:
                    if st.button("削除", key=f"del_fac_{staff}"):
                        config.remove_facility_staff(staff)
                        st.success(f"'{staff}' を削除しました。")
                        st.rerun()  # リスト更新のため必要

if __name__ == "__main__":
    main()