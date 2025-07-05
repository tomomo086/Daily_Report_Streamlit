import os
import tempfile
import shutil
from typing import Optional, Union, BinaryIO
from datetime import datetime
import secrets
import hashlib

class PrivacyFileHandler:
    """
    履歴に残らないファイル処理を行うクラス
    
    主な機能：
    - 一時ファイルの安全な作成・削除
    - メモリ上でのファイル処理
    - ファイル名の匿名化
    - 処理後の自動クリーンアップ
    """
    
    def __init__(self):
        self.temp_dir = None
        self.temp_files = []
        self.setup_temp_directory()
    
    def setup_temp_directory(self):
        """一時ディレクトリの設定"""
        try:
            # システムの一時ディレクトリを使用
            self.temp_dir = tempfile.mkdtemp(prefix='nippo_', suffix='_temp')
            # 一時ディレクトリのパーミッションを制限
            os.chmod(self.temp_dir, 0o700)
        except Exception as e:
            print(f"一時ディレクトリの作成に失敗: {e}")
    
    def generate_anonymous_filename(self, original_filename: Optional[str] = None) -> str:
        """
        匿名化されたファイル名を生成
        
        Args:
            original_filename: 元のファイル名（オプション）
        
        Returns:
            匿名化されたファイル名
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        random_suffix = secrets.token_hex(4)
        
        if original_filename:
            # 拡張子を保持
            _, ext = os.path.splitext(original_filename)
            return f"temp_{timestamp}_{random_suffix}{ext}"
        else:
            return f"temp_{timestamp}_{random_suffix}"
    
    def create_temp_file(self, content: bytes, filename: Optional[str] = None) -> str:
        """
        一時ファイルを作成
        
        Args:
            content: ファイルの内容
            filename: ファイル名（オプション）
        
        Returns:
            一時ファイルのパス
        """
        if not self.temp_dir:
            raise Exception("一時ディレクトリが設定されていません")
        
        if not filename:
            filename = self.generate_anonymous_filename()
        
        temp_file_path = os.path.join(self.temp_dir, filename)
        
        try:
            with open(temp_file_path, 'wb') as f:
                f.write(content)
            
            # ファイルのパーミッションを制限
            os.chmod(temp_file_path, 0o600)
            
            # 作成されたファイルをトラッキング
            self.temp_files.append(temp_file_path)
            
            return temp_file_path
        except Exception as e:
            print(f"一時ファイルの作成に失敗: {e}")
            raise
    
    def read_temp_file(self, file_path: str) -> bytes:
        """
        一時ファイルを読み取り
        
        Args:
            file_path: ファイルパス
        
        Returns:
            ファイルの内容
        """
        try:
            with open(file_path, 'rb') as f:
                return f.read()
        except Exception as e:
            print(f"ファイルの読み取りに失敗: {e}")
            raise
    
    def secure_delete_file(self, file_path: str):
        """
        ファイルを安全に削除（上書き削除）
        
        Args:
            file_path: 削除するファイルのパス
        """
        try:
            if os.path.exists(file_path):
                # ファイルサイズを取得
                file_size = os.path.getsize(file_path)
                
                # ファイルをランダムデータで上書き（3回）
                with open(file_path, 'r+b') as f:
                    for _ in range(3):
                        f.seek(0)
                        f.write(os.urandom(file_size))
                        f.flush()
                        os.fsync(f.fileno())
                
                # ファイルを削除
                os.remove(file_path)
                
                # トラッキングリストから削除
                if file_path in self.temp_files:
                    self.temp_files.remove(file_path)
                    
        except Exception as e:
            print(f"ファイルの安全削除に失敗: {e}")
    
    def cleanup_all(self):
        """すべての一時ファイルをクリーンアップ"""
        try:
            # 作成したすべての一時ファイルを安全に削除
            for file_path in self.temp_files.copy():
                self.secure_delete_file(file_path)
            
            # 一時ディレクトリを削除
            if self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
                
        except Exception as e:
            print(f"クリーンアップに失敗: {e}")
    
    def __enter__(self):
        """コンテキストマネージャー開始"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """コンテキストマネージャー終了時にクリーンアップ"""
        self.cleanup_all()

def get_anonymous_download_name(original_name: Optional[str] = None) -> str:
    """
    ダウンロード時の匿名ファイル名を生成
    
    Args:
        original_name: 元のファイル名
    
    Returns:
        匿名化されたファイル名
    """
    timestamp = datetime.now().strftime("%Y%m%d")
    random_id = secrets.token_hex(3)
    
    if original_name:
        name, ext = os.path.splitext(original_name)
        return f"report_{timestamp}_{random_id}{ext}"
    else:
        return f"report_{timestamp}_{random_id}.xlsx"

def clear_browser_cache_headers() -> dict:
    """
    ブラウザキャッシュを防ぐためのヘッダー
    
    Returns:
        キャッシュ防止ヘッダー
    """
    return {
        'Cache-Control': 'no-cache, no-store, must-revalidate, private',
        'Pragma': 'no-cache',
        'Expires': '0',
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY'
    }