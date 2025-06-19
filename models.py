from dataclasses import dataclass
from typing import List, Optional

@dataclass
class PatrolData:
    """巡回データを格納するクラス"""
    post4: str
    post5: str
    post1: str
    supervisor: str
    patrol_start: str
    large_theater_used: bool
    medium_theater_used: bool
    small_theater_used: bool
    weather: str  # 天気を追加
    
    @property
    def post4_lastname(self):
        return self.post4.split()[0] if self.post4 else ""
    
    @property
    def post5_lastname(self):
        return self.post5.split()[0] if self.post5 else ""
    
    @property
    def post1_lastname(self):
        return self.post1.split()[0] if self.post1 else ""
    
    @property
    def supervisor_lastname(self):
        return self.supervisor.split()[0] if self.supervisor else ""

@dataclass
class TimeRecord:
    """時間記録データ"""
    start_time: str
    end_time: str
    comment: str