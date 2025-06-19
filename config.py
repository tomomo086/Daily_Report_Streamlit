import json
import os
import streamlit as st

class Config:
    def __init__(self):
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_file = os.path.join(self.script_dir, "daily_report_config.json")
        self.security_staff_list = []
        self.facility_staff_list = []
        self.load()
    
    def load(self):
        """設定ファイルを読み込む"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.security_staff_list = data.get("security_staff", [])
                    self.facility_staff_list = data.get("facility_staff", [])
            except Exception as e:
                st.warning(f"設定ファイル読み込みエラー: {e}")
                # エラー時はデフォルト値を使用
                self.security_staff_list = []
                self.facility_staff_list = []
                st.info("デフォルト設定を使用します。")
    
    def save(self):
        """設定ファイルを保存する"""
        try:
            data = {
                "security_staff": self.security_staff_list,
                "facility_staff": self.facility_staff_list
            }
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            st.error(f"設定ファイルの保存に失敗しました: {e}\n保存先: {self.config_file}")
    
    def add_security_staff(self, name):
        """警備担当者を追加"""
        if name not in self.security_staff_list and name.strip():
            self.security_staff_list.append(name)
            self.save()
            return True
        return False
    
    def add_facility_staff(self, name):
        """設備担当者を追加"""
        if name not in self.facility_staff_list and name.strip():
            self.facility_staff_list.append(name)
            self.save()
            return True
        return False
    
    def remove_security_staff(self, name):
        """警備担当者を削除"""
        if name in self.security_staff_list:
            self.security_staff_list.remove(name)
            self.save()
            return True
        return False
    
    def remove_facility_staff(self, name):
        """設備担当者を削除"""
        if name in self.facility_staff_list:
            self.facility_staff_list.remove(name)
            self.save()
            return True
        return False