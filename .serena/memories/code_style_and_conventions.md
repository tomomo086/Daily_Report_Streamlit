# コードスタイルとコンベンション

## プログラミング言語
- **メイン言語**: Python 3.7+
- **文字エンコーディング**: UTF-8（日本語対応）

## コードスタイル

### 命名規則
- **クラス名**: PascalCase (`Config`, `PatrolData`, `ExcelWriter`)
- **関数・メソッド名**: snake_case (`add_security_staff`, `load_config`)
- **変数名**: snake_case (`security_staff_list`, `config_file`)
- **定数**: UPPER_SNAKE_CASE（使用例少ない）

### ファイル・ディレクトリ構造
- **メインファイル**: `streamlit_app.py` (エントリポイント)
- **モジュール**: 機能別ディレクトリ分割 (`excel/`, `utils/`)
- **設定**: `config.py` (クラスベース設定管理)
- **モデル**: `models.py` (データクラス定義)

### 日本語対応
- **コメント**: 日本語ドキュメント文字列使用
- **UI文字列**: すべて日本語表示
- **エラーメッセージ**: 日本語エラー表示
- **ファイル処理**: UTF-8エンコーディング明示

## 設計パターン

### クラス設計
```python
class Config:
    """設定管理クラスの典型例"""
    def __init__(self):
        # 初期化処理
    
    def load(self):
        """設定ファイルを読み込む"""
        
    def save(self):
        """設定ファイルを保存する"""
```

### エラーハンドリング
```python
try:
    # 処理実行
    with open(file_path, "r", encoding="utf-8") as f:
        data = json.load(f)
except Exception as e:
    st.warning(f"エラーメッセージ: {e}")
    # デフォルト値使用
```

### Streamlit UI パターン
```python
# タブ構造
tab1, tab2 = st.tabs(["📝 メイン機能", "👥 管理機能"])

# カラムレイアウト
col1, col2 = st.columns(2)

# セッション状態管理
if 'config' not in st.session_state:
    st.session_state.config = Config()
```

## データ管理

### JSON設定ファイル
- **ファイル名**: `daily_report_config.json`
- **エンコーディング**: UTF-8
- **フォーマット**: インデント2スペース、ASCII文字エスケープなし

### Excel処理
- **ライブラリ**: openpyxl（xlsxフォーマット対応）
- **結合セル対応**: 高度なセル操作実装
- **日本語対応**: UTF-8エンコーディング処理

## コードコメント方針
- **関数**: 日本語ドキュメント文字列
- **処理説明**: 必要に応じて行コメント
- **UI表示**: 日本語文字列リテラル使用