from datetime import datetime, timedelta
import random
from models import TimeRecord

class PatrolTimeGenerator:
    """巡回時間を生成するクラス"""
    
    def generate_4post_times(self, patrol_start, large, medium, small):
        """4ポストの巡回時間を生成"""
        if patrol_start == "22:00頃":
            return self._generate_4post_22(large, medium, small)
        else:  # 21:00頃 or その他
            return self._generate_4post_21(large, medium, small)
    
    def _generate_4post_21(self, large, medium, small):
        """21:00頃開始の4ポスト巡回時間"""
        start_time = datetime.strptime("21:00", "%H:%M") + timedelta(minutes=random.randint(0, 3))
        records = []
        
        # 巡回順序とコメントの設定
        patrol_order = [
            ('C17', 'E17', 10),
            ('C18', 'E18', 10),
            ('C23', 'E23', 10),
            ('C22', 'E22', 10),
            ('C16', 'E16', 10),
            ('C21', 'E21', 10),
            ('C15', 'E15', 10),
        ]
        
        comments = self._get_comments_21(large, medium, small)
        
        current_time = start_time
        for i, (start_cell, end_cell, duration) in enumerate(patrol_order):
            if i == 5 and medium and not large and not small:  # 中劇場のみ使用時
                records.append((start_cell, end_cell, 
                              TimeRecord("-", "-", comments[i])))
            elif i == 5 and medium:  # 中劇場使用時
                records.append((start_cell, end_cell, 
                              TimeRecord("-", "-", comments[i])))
            else:
                end_time = current_time + timedelta(minutes=duration)
                records.append((start_cell, end_cell, 
                              TimeRecord(current_time.strftime("%H:%M"), 
                                       end_time.strftime("%H:%M"), 
                                       comments[i])))
                current_time = end_time
        
        # 時間の調整（楽屋使用時）
        if medium:
            records[6] = ('C15', 'E15', 
                         TimeRecord(records[4][2].end_time, 
                                   (datetime.strptime(records[4][2].end_time, "%H:%M") + 
                                    timedelta(minutes=10)).strftime("%H:%M"), 
                                   comments[6]))
        
        return records
    
    def _generate_4post_22(self, large, medium, small):
        """22:00頃開始の4ポスト巡回時間"""
        start_time = datetime.strptime("22:00", "%H:%M") + timedelta(minutes=random.randint(0, 10))
        records = []
        
        patrol_order = [
            ('C17', 'E17', 5),
            ('C18', 'E18', 5),
            ('C23', 'E23', 5),
            ('C22', 'E22', 5),
            ('C16', 'E16', 5),
            ('C21', 'E21', 5),
            ('C15', 'E15', 5),
        ]
        
        comments = self._get_comments_22(large, medium, small)
        
        current_time = start_time
        for i, (start_cell, end_cell, duration) in enumerate(patrol_order):
            if i == 5 and medium:  # 中劇場使用時
                records.append((start_cell, end_cell, 
                              TimeRecord("-", "-", comments[i])))
            else:
                end_time = current_time + timedelta(minutes=duration)
                records.append((start_cell, end_cell, 
                              TimeRecord(current_time.strftime("%H:%M"), 
                                       end_time.strftime("%H:%M"), 
                                       comments[i])))
                current_time = end_time
        
        # 時間の調整（楽屋使用時）
        if medium:
            records[6] = ('C15', 'E15', 
                         TimeRecord(records[4][2].end_time, 
                                   (datetime.strptime(records[4][2].end_time, "%H:%M") + 
                                    timedelta(minutes=5)).strftime("%H:%M"), 
                                   comments[6]))
        
        return records
    
    def _get_comments_21(self, large, medium, small):
        """21:00頃開始時のコメント取得"""
        if large and medium and small:
            return [
                "異常なし",
                "異常なし",
                "異常なし",
                "異常なし(楽屋使用中の為その周辺は巡回実施せず)",
                "異常なし(楽屋使用中の為その周辺は巡回実施せず)",
                "楽屋使用中の為巡回実施せず",
                "異常なし(楽屋使用中の為その周辺は巡回実施せず)"
            ]
        elif large and medium:
            return [
                "異常なし",
                "異常なし",
                "異常なし",
                "異常なし(楽屋使用中の為その周辺は巡回実施せず)",
                "異常なし(楽屋使用中の為その周辺は巡回実施せず)",
                "楽屋使用中の為巡回実施せず",
                "異常なし(楽屋使用中の為その周辺は巡回実施せず)"
            ]
        elif large:
            return [
                "異常なし",
                "異常なし",
                "異常なし",
                "異常なし",
                "異常なし(楽屋使用中の為その周辺は巡回実施せず)",
                "異常なし",
                "異常なし(楽屋使用中の為その周辺は巡回実施せず)"
            ]
        elif medium:
            return [
                "異常なし",
                "異常なし",
                "異常なし",
                "異常なし(楽屋使用中の為その周辺は巡回実施せず)",
                "異常なし",
                "楽屋使用中の為巡回実施せず",
                "異常なし"
            ]
        else:
            return ["異常なし"] * 7
    
    def _get_comments_22(self, large, medium, small):
        """22:00頃開始時のコメント取得"""
        base = "異常なし(開始が遅いためトイレを重点的に巡回"
        if large and medium:
            return [
                base + ")",
                base + ")",
                base + ")",
                base + "、楽屋周りは除く)",
                base + "、楽屋周りは除く)",
                "楽屋使用中の為巡回実施せず",
                base + "、楽屋周りは除く)"
            ]
        elif large:
            return [
                base + ")",
                base + ")",
                base + ")",
                base + ")",
                base + "、楽屋周りは除く)",
                base + ")",
                base + "、楽屋周りは除く)"
            ]
        elif medium:
            return [
                base + ")",
                base + ")",
                base + ")",
                base + "、楽屋周りは除く)",
                base + ")",
                "楽屋使用中の為巡回実施せず",
                base + ")"
            ]
        else:
            return [base + ")"] * 7
    
    def generate_5post_times(self, large, medium, small):
        """5ポストの巡回時間を生成"""
        start_time = datetime.strptime("22:00", "%H:%M") + timedelta(minutes=random.randint(0, 5))
        
        records = []
        # 1回目の巡回
        end_time = start_time + timedelta(minutes=15)
        comment1 = self._get_5post_comment_1(large, medium, small)
        records.append(('C26', 'E26', 
                       TimeRecord(start_time.strftime("%H:%M"), 
                                 end_time.strftime("%H:%M"), 
                                 comment1)))
        
        # 2回目の巡回
        start_time2 = end_time
        end_time2 = start_time2 + timedelta(minutes=15)
        comment2 = self._get_5post_comment_2(large, medium, small)
        records.append(('C27', 'E27', 
                       TimeRecord(start_time2.strftime("%H:%M"), 
                                 end_time2.strftime("%H:%M"), 
                                 comment2)))
        
        return records
    
    def _get_5post_comment_1(self, large, medium, small):
        """5ポスト1回目のコメント"""
        if large:
            return "異常なし(楽屋使用中の為その周辺は巡回実施せず)"
        else:
            return "異常なし"
    
    def _get_5post_comment_2(self, large, medium, small):
        """5ポスト2回目のコメント"""
        if small or (large and medium and small):
            return "異常なし(楽屋使用中の為その周辺は巡回実施せず)"
        else:
            return "異常なし"
    
    def generate_other_times(self):
        """その他の時間を生成"""
        return {
            'morning_4post': (datetime.strptime("7:00", "%H:%M") + 
                            timedelta(minutes=random.randint(0, 20))).strftime("%H:%M"),
            'morning_5post': (datetime.strptime("7:15", "%H:%M") + 
                            timedelta(minutes=random.randint(0, 15))).strftime("%H:%M"),
            'morning_1post': (datetime.strptime("8:48", "%H:%M") + 
                            timedelta(minutes=random.randint(0, 8))).strftime("%H:%M"),
            'morning_4post_2': (datetime.strptime("7:30", "%H:%M") + 
                              timedelta(minutes=random.randint(0, 20))).strftime("%H:%M"),
            'night_4post': (datetime.strptime("22:00", "%H:%M") + 
                          timedelta(minutes=random.randint(0, 10))).strftime("%H:%M"),
            'patrol_4post': (datetime.strptime("21:30", "%H:%M") + 
                           timedelta(minutes=random.randint(0, 3))).strftime("%H:%M"),
            'patrol_4post_end': (datetime.strptime("22:50", "%H:%M") + 
                               timedelta(minutes=random.randint(0, 5))).strftime("%H:%M")
        }