# 技術スタックとアーキテクチャ

## 技術スタック

### フロントエンド
- **Streamlit 1.28+**: Web UI フレームワーク
- **Python 3.7+**: メインプログラミング言語

### バックエンド・データ処理
- **openpyxl 3.1.0+**: Excel操作（結合セル、フォーマット対応）
- **JSON**: 設定データ永続化

### 依存関係
最小限の依存関係で軽量構成:
```
streamlit>=1.28.0
openpyxl>=3.1.0
```

## プロジェクト構造

```
Daily_Report_Streamlit/
├── streamlit_app.py      # メインアプリケーション（エントリポイント）
├── config.py             # 設定管理（JSON永続化）
├── models.py             # データモデル（PatrolData, TimeRecord）
├── requirements.txt      # 依存関係定義
├── README.md            # プロジェクト詳細ドキュメント
├── 現実的な課題と学び.md    # 開発での学びと課題
├── images/              # アプリケーションスクリーンショット
│   ├── Screenshot.png   # メインUI画面
│   └── screenshot2.png  # 機能詳細画面
├── excel/               # Excel処理モジュール
│   ├── __init__.py
│   └── writer.py        # ExcelWriter クラス
└── utils/               # ユーティリティモジュール
    ├── __init__.py
    └── time_utils.py    # PatrolTimeGenerator クラス
```

## 主要コンポーネント

### Core Classes
- **Config**: JSON設定管理（スタッフリスト、設定永続化）
- **PatrolData**: 巡回データモデル
- **TimeRecord**: 時間記録モデル  
- **ExcelWriter**: Excel出力処理
- **PatrolTimeGenerator**: インテリジェント時間生成

### アーキテクチャパターン
- **MVC風設計**: UI(Streamlit) / Logic(Models) / Data(Config) 分離
- **モジュラー構造**: 機能別ディレクトリ分割
- **設定外部化**: JSON設定ファイルによる柔軟な運用