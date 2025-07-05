"""
Excelセル位置定義システム
各入力項目がExcelのどのセルに対応するかを定義
"""
from typing import Dict, List, Tuple, Any
from dataclasses import dataclass
from datetime import datetime

@dataclass
class CellDefinition:
    """セル定義クラス"""
    cell_address: str
    label: str
    data_type: str  # 'text', 'time', 'date', 'boolean', 'auto'
    required: bool = True
    auto_generate: bool = False
    description: str = ""
    
    def __post_init__(self):
        """自動生成の場合は必須項目をFalseに設定"""
        if self.auto_generate:
            self.required = False

class CellDefinitionManager:
    """セル定義管理クラス"""
    
    def __init__(self):
        self.definitions = self._initialize_definitions()
    
    def _initialize_definitions(self) -> Dict[str, CellDefinition]:
        """セル定義を初期化"""
        return {
            # 基本情報
            'weather': CellDefinition(
                'I4', '天気', 'text', True, False, 
                '天気情報（晴、曇、雨など）'
            ),
            'post4': CellDefinition(
                'F6', '4ポスト担当', 'text', True, False,
                '4ポスト担当者名'
            ),
            'post5': CellDefinition(
                'F7', '5ポスト担当', 'text', True, False,
                '5ポスト担当者名'
            ),
            'supervisor': CellDefinition(
                'J5', '設備担当者', 'text', True, False,
                '設備担当者名'
            ),
            'supervisor_2': CellDefinition(
                'J6', '設備担当者(2)', 'text', True, False,
                '設備担当者名（2箇所目）'
            ),
            
            # 勤務時間（勤務区分により自動設定）
            'work_time_1': CellDefinition(
                'K4', '勤務時間1', 'text', False, True,
                '勤務区分に応じた勤務時間'
            ),
            'work_time_2': CellDefinition(
                'L4', '勤務時間2', 'text', False, True,
                '勤務区分に応じた勤務時間'
            ),
            
            # 基本勤務時間
            'work_start_1': CellDefinition(
                'C10', '勤務開始時間1', 'time', False, True,
                '勤務区分に応じた開始時間'
            ),
            'work_start_2': CellDefinition(
                'D10', '勤務開始時間2', 'time', False, True,
                '勤務区分に応じた開始時間'
            ),
            'work_staff_1': CellDefinition(
                'E10', '勤務スタッフ1', 'text', False, True,
                '勤務区分に応じたスタッフ'
            ),
            'work_staff_2': CellDefinition(
                'F10', '勤務スタッフ2', 'text', False, True,
                '勤務区分に応じたスタッフ'
            ),
            'work_end': CellDefinition(
                'C11', '勤務終了時間', 'time', False, True,
                '勤務区分に応じた終了時間'
            ),
            'work_end_2': CellDefinition(
                'D11', '勤務終了時間2', 'time', False, True,
                '勤務区分に応じた終了時間'
            ),
            'work_end_staff': CellDefinition(
                'E11', '勤務終了スタッフ', 'text', False, True,
                '勤務区分に応じたスタッフ'
            ),
            
            # 4ポスト巡回記録（自動生成）
            'patrol_4post_1_start': CellDefinition(
                'C17', '4ポスト巡回1開始', 'time', False, True,
                '4ポスト巡回1開始時間'
            ),
            'patrol_4post_1_end': CellDefinition(
                'E17', '4ポスト巡回1終了', 'time', False, True,
                '4ポスト巡回1終了時間'
            ),
            'patrol_4post_1_comment': CellDefinition(
                'F17', '4ポスト巡回1コメント', 'text', False, True,
                '4ポスト巡回1のコメント'
            ),
            
            # 5ポスト巡回記録（自動生成）
            'patrol_5post_1_start': CellDefinition(
                'C26', '5ポスト巡回1開始', 'time', False, True,
                '5ポスト巡回1開始時間'
            ),
            'patrol_5post_1_end': CellDefinition(
                'E26', '5ポスト巡回1終了', 'time', False, True,
                '5ポスト巡回1終了時間'
            ),
            'patrol_5post_1_comment': CellDefinition(
                'F26', '5ポスト巡回1コメント', 'text', False, True,
                '5ポスト巡回1のコメント'
            ),
            
            # その他時間記録
            'morning_4post': CellDefinition(
                'E32', '朝4ポスト', 'time', False, True,
                '朝の4ポスト時間'
            ),
            'morning_5post': CellDefinition(
                'E34', '朝5ポスト', 'time', False, True,
                '朝の5ポスト時間'
            ),
            'morning_1post': CellDefinition(
                'E36', '朝1ポスト', 'time', False, True,
                '朝の1ポスト時間'
            ),
            'night_5post': CellDefinition(
                'H34', '夜5ポスト', 'time', False, True,
                '夜の5ポスト時間'
            ),
            'night_4post': CellDefinition(
                'H38', '夜4ポスト', 'time', False, True,
                '夜の4ポスト時間'
            ),
            'patrol_start_time': CellDefinition(
                'E41', '巡回開始', 'time', False, True,
                '巡回開始時間'
            ),
            'patrol_end_time': CellDefinition(
                'H41', '巡回終了', 'time', False, True,
                '巡回終了時間'
            ),
        }
    
    def get_user_input_cells(self) -> List[CellDefinition]:
        """ユーザー入力が必要なセルを取得"""
        return [cell for cell in self.definitions.values() 
                if cell.required and not cell.auto_generate]
    
    def get_auto_generate_cells(self) -> List[CellDefinition]:
        """自動生成されるセルを取得"""
        return [cell for cell in self.definitions.values() 
                if cell.auto_generate]
    
    def get_cell_by_address(self, address: str) -> CellDefinition:
        """セルアドレスから定義を取得"""
        for cell in self.definitions.values():
            if cell.cell_address == address:
                return cell
        raise ValueError(f"セルアドレス {address} が見つかりません")
    
    def get_cell_by_key(self, key: str) -> CellDefinition:
        """キーから定義を取得"""
        if key not in self.definitions:
            raise ValueError(f"キー {key} が見つかりません")
        return self.definitions[key]
    
    def get_all_cells_grouped(self) -> Dict[str, List[CellDefinition]]:
        """セルをグループ別に取得"""
        return {
            'user_input': self.get_user_input_cells(),
            'auto_generate': self.get_auto_generate_cells(),
            'all': list(self.definitions.values())
        }
    
    def validate_cell_value(self, key: str, value: Any) -> bool:
        """セルの値を検証"""
        cell = self.get_cell_by_key(key)
        
        if cell.required and not value:
            return False
        
        if cell.data_type == 'time' and value:
            try:
                datetime.strptime(str(value), '%H:%M')
                return True
            except ValueError:
                return False
        
        if cell.data_type == 'date' and value:
            try:
                datetime.strptime(str(value), '%Y-%m-%d')
                return True
            except ValueError:
                return False
        
        return True