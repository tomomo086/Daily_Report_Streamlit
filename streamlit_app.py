import streamlit as st
from datetime import datetime
import pandas as pd
from io import BytesIO
from models import PatrolData
from config import Config
from excel.writer import ExcelWriter
from excel.cell_definitions import CellDefinitionManager

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
    
    if 'cell_manager' not in st.session_state:
        st.session_state.cell_manager = CellDefinitionManager()
    
    config = st.session_state.config
    cell_manager = st.session_state.cell_manager
    
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
            supervisor = st.selectbox("設備担当者", supervisor_options, key="supervisor")

            # 天気入力欄
            st.subheader("天気")
            weather_options = ["晴", "曇", "雨", "晴/曇", "曇/雨", "その他"]
            weather_select = st.selectbox("天気を選択", weather_options, key="weather_select")
            weather = ""
            if weather_select == "その他":
                weather = st.text_input("天気を自由入力（上の選択肢以外の場合）", key="weather_input")
            else:
                weather = weather_select
            # 空の場合はデフォルト値
            if not weather:
                weather = "未記入"

            st.subheader("巡回設定")
            patrol_start = st.selectbox(
                "巡回開始時刻", 
                ["21:00頃", "22:00頃", "23:00頃"],
                key="patrol_start"
            )

            st.subheader("勤務区分")
            work_type = st.selectbox(
                "勤務区分を選択",
                ["通常", "早出", "残業"],
                key="work_type"
            )
        
        with col2:
            st.subheader("劇場使用状況")
            large_theater = st.checkbox("大劇場使用", key="large_theater")
            medium_theater = st.checkbox("中劇場（楽屋）使用", key="medium_theater")
            small_theater = st.checkbox("小劇場使用", key="small_theater")
            
            st.subheader("セル位置情報")
            
            # 今日の日付を表示
            today = datetime.today()
            st.info(f"📅 日付: {today.strftime('%Y年%m月%d日')} （自動入力）")
            
            # セル位置表示
            st.markdown("**📋 入力項目とセル位置:**")
            user_cells = cell_manager.get_user_input_cells()
            
            with st.expander("セル位置詳細", expanded=False):
                for cell in user_cells:
                    st.markdown(f"- **{cell.label}** → セル `{cell.cell_address}`")
                    if cell.description:
                        st.text(f"  {cell.description}")
            
            # 自動生成項目の表示
            st.markdown("**🤖 自動生成項目:**")
            auto_cells = cell_manager.get_auto_generate_cells()
            st.info(f"合計 {len(auto_cells)} 項目が自動生成されます")
            
            with st.expander("自動生成項目詳細", expanded=False):
                for cell in auto_cells:
                    st.markdown(f"- **{cell.label}** → セル `{cell.cell_address}`")
                    if cell.description:
                        st.text(f"  {cell.description}")
        
        st.markdown("---")
        
        if st.button("📋 日報作成", type="primary", use_container_width=True):
            if not all([post4, post5, post1, supervisor]):
                st.error("すべての担当者を選択してください。")
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
                            small_theater_used=small_theater,
                            weather=weather,
                            work_type=work_type
                        )
                        
                        writer = ExcelWriter()
                        
                        with st.spinner("日報を作成中..."):
                            output_bytes = writer.write_report(patrol_data)
                        
                        # output_bytesをセッションに保存
                        st.session_state['output_bytes'] = output_bytes
                        
                        today = datetime.today()
                        filename = f"日報_{today.strftime('%Y%m%d')}.xlsx"
                        
                        st.success("日報が正常に作成されました！")
                        
                        # ExcelファイルをWEB上で表示
                        st.subheader("📊 生成された日報")
                        
                        # 日報内容を表示ボタン
                        if st.button("📋 日報内容を表示", type="primary", use_container_width=True):
                            st.session_state.show_excel_content = True
                        
                        # Excelファイルの内容を表示
                        if st.session_state.get('show_excel_content', False):
                            st.markdown("---")
                            st.subheader("📋 日報内容")
                            
                            try:
                                # セッションからExcelファイルを取得
                                output_bytes = st.session_state.get('output_bytes', None)
                                if output_bytes is None:
                                    st.error("Excelファイルが見つかりません。先に日報を作成してください。")
                                else:
                                    excel_data = BytesIO(output_bytes)
                                    # シート名を取得
                                    xl = pd.ExcelFile(excel_data)
                                    if not xl.sheet_names:
                                        st.error("Excelファイルにシートが含まれていません。")
                                        st.info("生成されたファイルをご確認ください。")
                                    else:
                                        sheet_name = xl.sheet_names[0]
                                        excel_data.seek(0)
                                        df = pd.read_excel(excel_data, sheet_name=sheet_name, header=None)
                                        df = df.dropna(how='all').dropna(axis=1, how='all')
                                        st.markdown("**📊 Excelファイルの内容:**")
                                        column_config = {}
                                        for i in range(min(len(df.columns), 12)):
                                            column_letter = chr(65 + i)
                                            width = "medium" if i in [1, 5, 6, 9, 10, 11] else "small"
                                            column_config[i] = st.column_config.TextColumn(column_letter, width=width)
                                        st.dataframe(
                                            df,
                                            use_container_width=True,
                                            hide_index=True,
                                            column_config=column_config
                                        )
                                        st.info(f"📄 ファイル名: {filename} | 📅 シート名: {sheet_name}")
                                        if st.button("📋 内容を非表示", type="secondary"):
                                            st.session_state.show_excel_content = False
                                            st.rerun()
                            except Exception as e:
                                st.error(f"Excelファイルの表示中にエラーが発生しました: {e}")
                                st.info("生成されたファイルをご確認ください。")
                        
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
                        st.rerun()  # リフレッシュして入力をクリア
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
                        st.rerun()  # リフレッシュして入力をクリア
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
