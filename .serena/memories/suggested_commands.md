# 推奨コマンド一覧

## 必須コマンド

### アプリケーション実行
```bash
# メインコマンド - Streamlitアプリ起動
streamlit run streamlit_app.py

# 開発モード - 自動リロード有効
streamlit run streamlit_app.py --server.runOnSave true
```

### 初期セットアップ
```bash
# プロジェクトディレクトリ移動
cd "C:\Users\tomon\dev\projects\Daily_Report_Streamlit"

# 依存関係インストール
pip install -r requirements.txt
```

## 開発支援コマンド

### 依存関係管理
```bash
# 依存関係確認
pip list | findstr -E "(streamlit|openpyxl)"

# 個別インストール
pip install streamlit>=1.28.0
pip install openpyxl>=3.1.0
```

### トラブルシューティング
```bash
# Streamlitキャッシュクリア
streamlit cache clear

# 設定ファイル確認
type daily_report_config.json

# プロセス確認
tasklist | findstr streamlit
netstat -an | findstr :8501
```

## ファイル操作（Windows）

### 基本操作
```cmd
# ファイル一覧
dir
dir /s *.py

# ファイル検索
findstr /s "検索文字列" *.py

# ファイル表示
type ファイル名.py
more ファイル名.py
```

### Git操作（必要に応じて）
```bash
git status
git add .
git commit -m "変更内容"
```

## 実行時の注意事項

### ポート設定
- デフォルトポート: 8501
- カスタムポート: `streamlit run streamlit_app.py --server.port 8502`

### ブラウザアクセス
- ローカル: http://localhost:8501
- ネットワーク: http://[IPアドレス]:8501

### ファイルアップロード制限
- 最大ファイルサイズ: 10MB（アプリ内制限）
- 対応ファイル形式: .xlsx（Excelファイルのみ）

## 開発フロー推奨コマンド実行順序

1. `cd "C:\Users\tomon\dev\projects\Daily_Report_Streamlit"`
2. `pip install -r requirements.txt` (初回のみ)
3. `streamlit run streamlit_app.py`
4. ブラウザで http://localhost:8501 にアクセス
5. 開発中は Ctrl+C でサーバー停止、再実行