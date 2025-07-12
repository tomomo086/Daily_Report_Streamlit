# 📋 日報作成ツール (Streamlit版)

警備業務の日報を効率的に作成するWebアプリケーションです。Excelテンプレートに基づいて、担当者情報、巡回記録、天気などを自動入力し、完成した日報をダウンロードできます。

## 🚀 機能一覧

### 📝 日報作成機能
- **担当者選択**: 4ポスト、5ポスト、1ポスト、設備担当者の選択
- **天気入力**: プリセット選択または自由入力
- **巡回設定**: 開始時刻の選択（21:00頃/22:00頃）
- **勤務区分**: 通常/早出/残業の選択
- **施設使用状況**: 大劇場、中劇場（楽屋）、小劇場の使用状況
- **Excelテンプレート処理**: アップロードされたテンプレートへの自動入力
- **時間自動生成**: 巡回時間のランダム生成とリアルな記録作成

### 👥 スタッフ管理機能
- **警備担当者管理**: 担当者の追加・削除
- **設備担当者管理**: 設備担当者の追加・削除
- **データ永続化**: JSON形式での設定保存

### 🔧 高度な機能
- **担当者重複チェック**: 同一人物の複数ポスト割り当て防止
- **時間計算ロジック**: 使用状況に応じた巡回時間とコメントの自動調整
- **Excelセル操作**: 結合セル対応、フォント設定、安全な値設定
- **エラーハンドリング**: 堅牢なファイル処理とエラー回復

## 🛠️ 技術スタック

- **フロントエンド**: Streamlit
- **バックエンド**: Python 3.7+
- **Excel処理**: openpyxl
- **データ管理**: JSON (設定ファイル)

## 📦 インストール

### 必要な環境
- Python 3.7以上
- pip

### セットアップ手順

1. **リポジトリのクローン**
```bash
git clone https://github.com/yourusername/daily-report-tool.git
cd daily-report-tool
```

2. **仮想環境の作成（推奨）**
```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
```

3. **依存関係のインストール**
```bash
pip install -r requirements.txt
```

4. **アプリケーションの起動**
```bash
streamlit run streamlit_app.py
```

## 🚀 デプロイ

### Streamlit Community Cloud
1. GitHubのプライベートリポジトリにコードをプッシュ
2. [Streamlit Community Cloud](https://streamlit.io/cloud)でアカウント作成
3. プライベートアクセス権限を設定
4. リポジトリを選択してデプロイ

### Railway
```bash
# Railway CLIを使用
railway login
railway init
railway up
```

## 📖 使用方法

### 基本的な日報作成フロー

1. **スタッフ管理タブでの事前設定**
   - 警備担当者を登録
   - 設備担当者を登録

2. **日報作成タブでの入力**
   - 各ポストの担当者を選択
   - 天気を選択または入力
   - 巡回開始時刻を選択
   - 勤務区分を選択
   - 劇場使用状況をチェック
   - Excelテンプレートファイルをアップロード

3. **日報作成ボタンをクリック**
   - 自動的に時間計算と記録生成
   - Excelファイルの完成
   - ダウンロードボタンから取得

### ファイル形式について

- **入力**: `.xlsx`形式のExcelテンプレート（最大10MB）
- **出力**: `日報_YYYYMMDD.xlsx`形式

## 🏗️ プロジェクト構造

```
daily-report-tool/
├── streamlit_app.py      # メインアプリケーション
├── config.py             # 設定管理クラス
├── models.py             # データモデル定義
├── requirements.txt      # 依存関係
├── .gitignore           # Git除外設定
├── README.md            # このファイル
├── excel/               # Excel処理モジュール
│   ├── __init__.py
│   └── writer.py        # Excel書き込み処理
└── utils/               # ユーティリティ
    ├── __init__.py
    └── time_utils.py    # 時間生成ロジック
```

## ⚙️ 設定ファイル

アプリケーションは`daily_report_config.json`に設定を保存します：

```json
{
  "security_staff": ["担当者1", "担当者2"],
  "facility_staff": ["設備担当1", "設備担当2"]
}
```

## 🔧 カスタマイズ

### 巡回時間の調整
`utils/time_utils.py`の`PatrolTimeGenerator`クラスで時間生成ロジックをカスタマイズできます：

```python
# 例: 基準時間の変更
start_time = datetime.strptime("21:00", "%H:%M") + timedelta(minutes=random.randint(0, 3))
```

### コメントのカスタマイズ
使用状況に応じたコメントは`_get_comments_21()`、`_get_comments_22()`メソッドで設定されています。

## 🐛 トラブルシューティング

### よくある問題

1. **Excelファイルが読み込めない**
   - ファイル形式が`.xlsx`であることを確認
   - ファイルサイズが10MB以下であることを確認
   - テンプレートに必要なシートが存在することを確認

2. **設定が保存されない**
   - アプリケーション実行ディレクトリの書き込み権限を確認
   - `daily_report_config.json`ファイルの権限を確認

3. **時間生成がおかしい**
   - システム時刻が正しく設定されていることを確認

## 🤝 開発への貢献

1. このリポジトリをフォーク
2. フィーチャーブランチを作成 (`git checkout -b feature/amazing-feature`)
3. 変更をコミット (`git commit -m 'Add amazing feature'`)
4. ブランチにプッシュ (`git push origin feature/amazing-feature`)
5. プルリクエストを作成

## 📄 ライセンス(非公開)



## 🙏 謝辞

- [Streamlit](https://streamlit.io/) - 素晴らしいWebアプリフレームワーク
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel処理ライブラリ

## 📞 サポート

問題や質問がある場合は、[Issues](https://github.com/yourusername/daily-report-tool/issues)ページで報告してください。

---

**開発者**: tomomo086 ＋ Claude Sonnet4
**バージョン**: 1.0.0  
**最終更新**: 2025年6月27日
