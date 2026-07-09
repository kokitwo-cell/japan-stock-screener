# 日本株スクリーナー

東証プライム・スタンダード・グロース 約3,900社を対象としたWebスクリーニングツール。

🌐 **ライブサイト**: https://kokitwo-cell.github.io/japan-stock-screener/

## 概要

業績トレンド（売上高・営業利益・EPSの右肩上がり）、配当の継続性（減配なし・連続増配）、配当利回り、PER/PBR、東証33業種等で約3,900社を絞り込めます。データはGitHub上に静的JSONとして保管され、ブラウザだけで完結します。

## ポートフォリオ管理

ヘッダーの「📊 ポートフォリオ」タブから保有資産を確認できます。

- **保有一覧**: `data/portfolio.json` を初期データとして読み込み、株価・配当・業種はスクリーナーのキャッシュ（毎日自動更新）から反映。評価額・損益・年間配当（NISA株数を考慮した税引後）・利回り・業種構成を表示
- **編集**: 株数・NISA株数・取得単価はテーブル上で直接編集でき、銘柄の追加・削除も可能。変更はブラウザ（localStorage）に保存され、「初期データに戻す」でリセット
- **組入シミュレーション**: 銘柄と株数を指定すると「現在 vs 組入後」で評価額・年間配当・利回り・業種構成の変化を並べて表示。スクリーナーの銘柄詳細モーダルの「📊 組入試算」からもワンクリックで呼び出せます

保有内容が変わったら `data/portfolio.json` を更新してコミットすると、どのブラウザでも初期データとして反映されます。

## ファイル構成

```
.
├── index.html                       # フロント（スクリーニング + ポートフォリオ管理）
├── data/
│   ├── portfolio.json               # 保有銘柄の初期データ（コード・株数・取得単価・NISA株数）
│   ├── stock_cache.json             # 銘柄データ本体（業績・配当・株価等）
│   ├── tse_codes.json               # 東証銘柄コード一覧（JPX由来）
│   └── jquants_info.json            # 日本語名・東証33業種（J-Quants由来）
├── scripts/
│   ├── update_data.py               # データ更新スクリプト
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
