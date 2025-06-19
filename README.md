# 日報作成ツール (Streamlit版)

プロジェクトの日報作成を自動化するWebアプリケーションです。

## 特徴

- **Webブラウザで使用可能** - クラウドでの利用に最適
- **担当者管理** - 担当者の登録・管理
- **自動時間生成** - 複雑な巡回時間計算を自動化
- **Excel出力** - 既存のテンプレートに自動入力

## プロジェクト構造

```
Daily_Report_Streamlit/
├── streamlit_daily_report.py    # メインアプリケーション
├── models.py                    # データモデル
├── config.py                    # 設定管理
├── utils/
│   ├── __init__.py
│   └── time_utils.py           # 時間生成ロジック
├── excel/
│   ├── __init__.py
│   └── writer.py               # Excel書き込み処理
├── requirements.txt            # 依存関係
└── README.md                   # このファイル
```

## インストール・使用方法

1. 依存関係のインストール:
```bash
pip install -r requirements.txt
```

2. アプリケーション起動:
```bash
streamlit run streamlit_daily_report.py
```

3. ブラウザで http://localhost:8501 にアクセス

## 機能

### 日報作成タブ
- 担当者選択
- 巡回開始時刻選択（21:00頃/22:00頃）
- 劇場使用状況設定
- Excelテンプレートアップロード
- 自動生成された日報のダウンロード

### スタッフ管理タブ
- 警備担当者の追加・削除
- 設備担当者の追加・削除
- 設定の自動保存（daily_report_config.json）

## 元プロジェクトとの違い

- **GUI**: tkinter → Streamlit（Web UI）
- **配布**: デスクトップアプリ → Webアプリ（クラウド対応）
- **ファイル構造**: 元の構造を保持しつつ、Streamlit用に最適化

## 保守性の改善

- モジュール分割により各機能が独立
- 既存のロジックを再利用
- テストやデバッグが容易
- 新機能追加時の影響範囲を限定