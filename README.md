# 日本株スクリーナー

東証プライム・スタンダード・グロース 約3,900社を対象としたWebスクリーニングツール。

🌐 **ライブサイト**: https://kokitwo-cell.github.io/japan-stock-screener/

## 概要

業績トレンド（売上高・営業利益・EPSの右肩上がり）、配当の継続性（減配なし・連続増配）、配当利回り、PER/PBR、東証33業種等で約3,900社を絞り込めます。データはGitHub上に静的JSONとして保管され、ブラウザだけで完結します。

## ポートフォリオ管理

ヘッダーの「📊 ポートフォリオ」タブから保有資産を確認できます。

- **保有一覧**: `data/portfolio.json` を初期データとして読み込み、株価・配当・業種はスクリーナーのキャッシュ（毎日自動更新）から反映。評価額・損益・年間配当（NISA株数を考慮した税引後）・利回り・業種構成を表示
- **編集**: 株数・NISA株数・取得単価はテーブル上で直接編集でき、銘柄の追加・削除も可能。変更はブラウザ（localStorage）に保存され、「初期データに戻す」でリセット
- **組入・売却シミュレーション**: 複数銘柄の組入（株価も手動指定可）と売却を同時に指定でき、「現在 vs 組入後」で評価額・年間配当・利回り・業種構成・ディフェンシブ比率の変化と入金/出金/差引資金を表示
- **詳細モーダル**: 保有一覧の銘柄名クリックでスクリーナーと同じ業績・配当チャートを表示
- **資産推移**: 平日の株価更新ごとに評価額を `data/history.json` に記録し、推移グラフを自動描画
- **保有中バッジ / 一時配当の警告**: スクリーナーで保有銘柄に「保有中」バッジを表示。最新配当が前年比で急増している銘柄（特別・記念配当の可能性）は利回りに ⚠ を付与
- **エクスポート**: 「⧉ データをコピー」「↓ JSON保存」で現在の保有を `data/portfolio.json` 形式で書き出せます

保有内容が変わったら `data/portfolio.json` を更新してコミットすると、どのブラウザでも初期データとして反映されます（サイドバーの「⧉ データをコピー」で現在の状態を貼り替え用に取得できます）。

## ファイル構成

```
.
├── index.html                       # フロント（スクリーニング + ポートフォリオ管理）
├── data/
│   ├── portfolio.json               # 保有銘柄の初期データ（コード・株数・取得単価・NISA株数）
│   ├── history.json                 # 資産推移の日次スナップショット（自動追記）
│   ├── stock_cache.json             # 銘柄データ本体（業績・配当・株価等）
│   ├── tse_codes.json               # 東証銘柄コード一覧（JPX由来）
│   └── jquants_info.json            # 日本語名・東証33業種（J-Quants由来）
├── scripts/
│   ├── update_data.py               # データ更新スクリプト
│   ├── record_history.py            # ポートフォリオ評価額を history.json に日次記録
│   └── requirements.txt
└── .github/workflows/
    ├── deploy-pages.yml             # GitHub Pages 自動デプロイ
    └── update-data.yml              # データ更新（週次cron + 手動実行）
```

## データ更新

### 自動（推奨）
GitHub Actions が **毎週日曜 00:00 UTC（月曜 09:00 JST）** に株価と配当利回りを再計算します。

### 手動
GitHubのActionsタブ → 「Update stock data」 → Run workflow から以下を選択：
- `prices`  ... 株価のみ更新（数十分）
- `full`    ... 全銘柄をyfinanceから取り直し（数時間、夜間推奨）
- `irbank`  ... ir-bankからの長期業績補完（数時間）

### ローカル
```bash
pip install -r scripts/requirements.txt

# 株価のみ更新
UPDATE_PRICES_ONLY=1 python scripts/update_data.py

# 全銘柄取得
python scripts/update_data.py

# ir-bank補完
ENRICH_IRBANK=1 python scripts/update_data.py
```

## J-Quants APIキー（任意）
日本語名・東証33業種を最新化したい場合は、リポジトリのSettings → Secrets and variables → Actions に `JQUANTS_REFRESH_TOKEN` を登録。
