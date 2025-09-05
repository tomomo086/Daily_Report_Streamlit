# 開発コマンドと運用手順

## 実行コマンド

### アプリケーション起動
```bash
# Streamlitアプリ実行（メインコマンド）
streamlit run streamlit_app.py

# ローカル開発サーバー起動（自動リロード有効）
streamlit run streamlit_app.py --server.runOnSave true
```

### 依存関係管理
```bash
# 依存関係インストール
pip install -r requirements.txt

# 個別パッケージインストール
pip install streamlit>=1.28.0
pip install openpyxl>=3.1.0

# 依存関係確認
pip list | grep -E "(streamlit|openpyxl)"
```

## Windows システムコマンド

### ファイル操作
```cmd
# ディレクトリ一覧
dir
ls  # PowerShellの場合

# ファイル検索
findstr /s "検索文字列" *.py
dir /s *.py | findstr "ファイル名"

# ファイルコピー
copy source_file target_file
xcopy /s source_dir target_dir

# ディレクトリ移動
cd "フルパス"
cd ..
```

### プロセス管理
```cmd
# プロセス確認
tasklist | findstr python
tasklist | findstr streamlit

# ポート使用確認
netstat -an | findstr :8501
```

## 開発ワークフロー

### 新機能開発手順
1. `models.py` でデータ構造更新
2. `config.py` で設定項目追加
3. `streamlit_app.py` でUI実装
4. `excel/writer.py` でExcel出力対応
5. 動作確認とテスト実行

### トラブルシューティング
```bash
# Streamlit キャッシュクリア
streamlit cache clear

# 設定ファイル確認
type daily_report_config.json  # Windows
cat daily_report_config.json   # PowerShell/Git Bash

# ログ確認
# Streamlitは標準出力にログ表示
```

## ファイル管理

### 設定ファイル
- `daily_report_config.json`: スタッフリスト設定（自動生成）
- `requirements.txt`: Python依存関係定義

### 出力ファイル
- アプリ内でアップロードされたExcelファイルを処理し、ダウンロード可能な日報Excel生成