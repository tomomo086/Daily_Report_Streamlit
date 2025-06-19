from openpyxl import load_workbook
from openpyxl.styles import Font
from datetime import datetime
import io
import streamlit as st
from models import PatrolData
from utils.time_utils import PatrolTimeGenerator

class ExcelWriter:
    def __init__(self):
        self.time_generator = PatrolTimeGenerator()
    
    def write_report(self, file_bytes, patrol_data: PatrolData):
        """日報をExcelファイルに書き込む"""
        wb = load_workbook(io.BytesIO(file_bytes))
        
        # シート名の取得
        today = datetime.today()
        sheet_name = f"{today.month}.{today.day}"
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"シート {sheet_name} が見つかりません。")
        
        ws = wb[sheet_name]
        
        # 基本情報の書き込み
        self._write_basic_info(ws, patrol_data)
        
        # 巡回記録の書き込み
        self._write_patrol_records(ws, patrol_data)
        
        # その他の時間記録
        self._write_other_records(ws, patrol_data)
        
        # フォントサイズの設定
        self._set_font_sizes(wb)
        
        # バイト配列として返す
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()
    
    def _safe_set_cell_value(self, ws, cell_address, value):
        """結合セルかどうかをチェックしてから値を設定"""
        try:
            cell = ws[cell_address]
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    top_left = merged_range.start_cell
                    ws[top_left.coordinate] = value
                    return
            cell.value = value
        except Exception as e:
            st.warning(f"セル {cell_address} への値設定でエラー: {e}")
            # エラーが発生しても処理を継続
            try:
                ws[cell_address] = value
            except Exception as fallback_error:
                st.error(f"フォールバック処理でもエラー: {fallback_error}")
    
    def _set_time(self, ws, cell, time_str):
        """時間をセルに設定"""
        if time_str:
            try:
                if isinstance(time_str, str):
                    # time型で格納
                    time_obj = datetime.strptime(time_str, "%H:%M").time()
                else:
                    time_obj = time_str
                self._safe_set_cell_value(ws, cell, time_obj)
                try:
                    ws[cell].number_format = 'hh:mm'
                except Exception as format_error:
                    st.write(f"時間フォーマット設定エラー: {format_error}")
            except ValueError as time_error:
                self._safe_set_cell_value(ws, cell, time_str)
    
    def _write_basic_info(self, ws, patrol_data: PatrolData):
        """基本情報を書き込む"""
        self._safe_set_cell_value(ws, 'I4', patrol_data.weather)
        self._safe_set_cell_value(ws, 'F6', patrol_data.post4)
        self._safe_set_cell_value(ws, 'F7', patrol_data.post5)
        # 設備担当者は登録された通りに出力（そのまま）
        self._safe_set_cell_value(ws, 'J5', patrol_data.supervisor)
        self._safe_set_cell_value(ws, 'J6', patrol_data.supervisor)

        # 勤務区分による分岐
        if getattr(patrol_data, 'work_type', '通常') == '早出':
            # 早出
            self._safe_set_cell_value(ws, 'K4', '7:30～23:00')
            self._safe_set_cell_value(ws, 'L4', '7:30～23:00')
            self._set_time(ws, 'C10', '7:30')
            self._set_time(ws, 'D10', '7:30')
            self._safe_set_cell_value(ws, 'E10', patrol_data.post4)
            self._safe_set_cell_value(ws, 'F10', patrol_data.post4)
            self._set_time(ws, 'C11', '23:00')
            self._safe_set_cell_value(ws, 'E11', patrol_data.post4)
        elif getattr(patrol_data, 'work_type', '通常') == '残業':
            # 残業
            self._safe_set_cell_value(ws, 'K4', '8:00～24:00')
            self._safe_set_cell_value(ws, 'L4', '8:00～24:00')
            self._set_time(ws, 'C10', '8:00')
            self._safe_set_cell_value(ws, 'E10', patrol_data.post1)
            self._set_time(ws, 'C11', '24:00')
            self._set_time(ws, 'D11', '24:00')
            self._safe_set_cell_value(ws, 'E11', patrol_data.post4)
        else:
            # 通常
            self._set_time(ws, 'C10', '8:00')
            self._safe_set_cell_value(ws, 'E10', patrol_data.post1)
            self._set_time(ws, 'C11', '23:00')
            self._safe_set_cell_value(ws, 'E11', patrol_data.post4)
    
    def _write_patrol_records(self, ws, patrol_data: PatrolData):
        """巡回記録を書き込む"""
        # 4ポストの巡回
        records = self.time_generator.generate_4post_times(
            patrol_data.patrol_start,
            patrol_data.large_theater_used,
            patrol_data.medium_theater_used,
            patrol_data.small_theater_used
        )
        
        for cell_start, cell_end, record in records:
            if record.start_time != "-":
                self._set_time(ws, cell_start, record.start_time)
            else:
                self._safe_set_cell_value(ws, cell_start, "-")
                
            if record.end_time != "-":
                self._set_time(ws, cell_end, record.end_time)
            else:
                self._safe_set_cell_value(ws, cell_end, "-")
            
            # コメントの書き込み
            comment_cell = cell_start.replace('C', 'F').replace('E', 'F')
            self._safe_set_cell_value(ws, comment_cell, record.comment)
        
        # 5ポストの巡回
        records_5post = self.time_generator.generate_5post_times(
            patrol_data.large_theater_used,
            patrol_data.medium_theater_used,
            patrol_data.small_theater_used
        )
        
        for cell_start, cell_end, record in records_5post:
            self._set_time(ws, cell_start, record.start_time)
            self._set_time(ws, cell_end, record.end_time)
            comment_cell = cell_start.replace('C', 'F')
            self._safe_set_cell_value(ws, comment_cell, record.comment)
    
    def _write_other_records(self, ws, patrol_data: PatrolData):
        """その他の時間記録を書き込む"""
        other_times = self.time_generator.generate_other_times()
        
        for cell, time, name_attr in [
            ('E32', other_times['morning_4post'], 'post5_lastname'),
            ('E34', other_times['morning_5post'], 'post5_lastname'),
            ('E36', other_times['morning_1post'], 'post1_lastname'),
            ('E38', other_times['morning_4post_2'], 'post4_lastname')
        ]:
            self._safe_set_cell_value(ws, cell, time.lstrip("0"))
            name_cell = cell.replace('E', 'G')
            self._safe_set_cell_value(ws, name_cell, 
                                    getattr(patrol_data, name_attr))
        
        # 夜の記録
        self._set_time(ws, 'H34', "22:50")
        self._safe_set_cell_value(ws, 'J34', patrol_data.post5_lastname)
        
        self._set_time(ws, 'H38', other_times['night_4post'])
        self._safe_set_cell_value(ws, 'J38', patrol_data.post4_lastname)
        
        self._set_time(ws, 'E41', other_times['patrol_4post'])
        self._safe_set_cell_value(ws, 'G41', patrol_data.post4_lastname)
        
        self._set_time(ws, 'H41', other_times['patrol_4post_end'])
        self._safe_set_cell_value(ws, 'J41', patrol_data.post4_lastname)
    
    def _set_font_sizes(self, wb):
        """特定セルのフォントサイズを設定"""
        for ws_item in wb.worksheets:
            for cell in ['I5', 'I6', 'K5', 'K6']:
                try:
                    ws_item[cell].font = Font(size=8)
                except Exception as font_error:
                    st.write(f"フォント設定エラー (セル {cell}): {font_error}")

# 追加ボタンが押されたら、rerun で初期値に戻す
st.session_state.new_security = ""
st.session_state.new_facility = ""
new_security = st.text_input("新しい警備担当者を追加", key="new_security")
if st.button("追加", key="add_security"):
    if new_security:
        if config.add_security_staff(new_security):
            st.success(f"'{new_security}' を追加しました。")
            st.experimental_rerun()  # これでフィールドが初期化される
        else:
            st.warning("既に登録されているか、無効な名前です。")
    else:
        st.warning("名前を入力してください。")