# タスク完了時のワークフロー

## コード変更後の基本チェック

### 1. 構文・実行確認
```bash
# Python構文チェック
python -m py_compile streamlit_app.py
python -m py_compile config.py
python -m py_compile models.py

# アプリケーション起動確認
streamlit run streamlit_app.py
# ブラウザで動作確認 → Ctrl+C で停止
```

### 2. 機能テスト手順

#### UI機能確認
1. スタッフ管理タブでの担当者追加・削除
2. 日報作成タブでの各入力項目動作
3. Excelファイルアップロード機能
4. 日報生成ボタンの動作

#### データ永続化確認
```bash
# 設定ファイル確認
type daily_report_config.json

# 設定内容の正常性確認
python -c "import json; print(json.load(open('daily_report_config.json', 'r', encoding='utf-8')))"
```

## 品質保証プロセス

### コードレビューポイント
1. **日本語エンコーディング**: UTF-8対応確認
2. **エラーハンドリング**: try-except文での適切な例外処理
3. **Streamlit UI**: レスポンシブ表示とユーザビリティ
4. **Excel処理**: openpyxlでの結合セル対応

### セキュリティチェック
1. **ファイルアップロード**: サイズ制限（10MB）とファイル形式制限
2. **データ入力**: 適切な入力検証とサニタイゼーション
3. **設定ファイル**: 機密情報の非含有確認

## リリース前最終確認

### 動作環境テスト
```bash
# 依存関係確認
pip list | findstr -E "(streamlit|openpyxl)"

# クリーンな環境での起動確認
streamlit cache clear
streamlit run streamlit_app.py
```

### ドキュメント更新
1. **README.md**: 機能変更時の説明更新
2. **requirements.txt**: 新規依存関係追加時の更新
3. **コメント**: 新機能の日本語ドキュメント文字列追加

## トラブルシューティング対応

### 一般的な問題と対処
```bash
# ポート競合時
netstat -an | findstr :8501
taskkill /f /pid [プロセスID]

# 依存関係問題時
pip install --force-reinstall streamlit openpyxl

# キャッシュ問題時
streamlit cache clear
```

### ログ確認
- Streamlitの標準出力でエラー確認
- `st.error()`, `st.warning()`, `st.info()` でのユーザー向けメッセージ

## デプロイメント準備（必要に応じて）

### ファイル整理
1. 不要なファイル削除
2. `.gitignore` 確認
3. `daily_report_config.json` の初期化または削除

### 環境確認
```bash
# Python バージョン確認
python --version

# 最小構成での動作確認
pip freeze > current_requirements.txt
```

## 完了基準
✅ 構文エラーなし  
✅ UI動作正常  
✅ データ保存・読み込み正常  
✅ Excel処理正常  
✅ エラーハンドリング適切  
✅ 日本語表示正常  
✅ ドキュメント更新完了