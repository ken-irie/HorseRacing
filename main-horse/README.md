# main_horse_decide.py

netkeiba.com から WIN5 馬券の対象レース情報をスクレイピングし、各レースの馬柱データ（過去5走まで）を抽出して Excel ファイルにエクスポートするツールです。

## 概要

このスクリプトは以下の処理を自動化します：

1. **WIN5 対象レース取得** - netkeiba.com の WIN5 ページから 対象レースの race_id（5つまで）と開催日を抽出
2. **馬柱データ取得** - 各レースの馬柱ページから馬番、馬名、性齢、騎手名を取得
3. **過去走データ抽出** - 各馬の過去4走（前走、2走、3走、4走）の詳細情報をパース
4. **Excel エクスポート** - データを新規 Excel ファイルに書き込み、自動的に格子罫線を適用

## 主要機能

### WIN5 レース情報取得
- `pick_win5_ids()` - WIN5 ページから対象レースの race_id と開催日（YYYYMMDD形式）を取得
- `_extract_ids_from_html()` - HTMLから race_id を正規表現でパース
- `_race_date()` - 開催日を抽出（年月日を自動判定）

### 馬柱データ抽出
- `extract_horse_table()` - 馬柱テーブルから全馬の情報をデータフレームに変換
- `parse_past_cell()` - 過去1走分のセルから詳細情報をパース
- `fetch_shutsuba_with_meta()` - レースメタ情報（日付、レース名、開催地、レース番号）を取得

### 抽出データ項目

**基本情報**
- 馬番（枠番）
- 馬名
- 性齢（性別と年齢）
- 騎手名

**過去走データ（前走/2走/3走/4走）**
- レース名
- 開催地（札幌、東京、中山など）
- コース（芝1600、ダート1400など）
- 着順（1着、2着など）
- 着差（秒数、ハナ長さなど）
- 通過順（4-3-4-3形式）
- 3F タイム（最後の3ハロン＝600m のタイム）

## 使用方法

### 基本的な実行

```bash
python main_horse_decide.py
```

土曜日（idx=0）の WIN5 デフォルトレースを処理します。

### WIN5 ページ URL を指定

```bash
python main_horse_decide.py "https://race.netkeiba.com/top/win5.html?idx=1"
```

日曜日（idx=1）の WIN5 を処理したい場合など、URL を指定できます。

## 入出力

### 入力
- **WIN5 ページ URL**（オプション）
  - デフォルト: `https://race.netkeiba.com/top/win5.html?idx=0`（土曜日）
  - 環境変数 `PC_URL` または コマンドラインで上書き可能
  - 対象は最大 5 レース

### 出力
- **ファイル名**: `output/Win5軸馬決定_YYYYMMDD_TIMESTAMP.xlsx`
  - YYYYMMDD: 開催日
  - TIMESTAMP: スクリプト実行日時（YYYYMMDDHHmmss）
- **シート構成**: 各レース 1 枚のシート
  - シート名は「開催地_レース番号_レース名」（例：`京都_9R_秋明菊賞`）
  - 行：各馬（最大18頭）
  - 列：馬番、馬名、性齢、騎手名、過去走データ項目

## 依存パッケージ

```
requests          # HTTP 通信
beautifulsoup4    # HTML パース
pandas            # データフレーム処理
openpyxl          # Excel ファイル作成
selenium          # ブラウザ自動化（インポートのみ、このスクリプトではは使用していません）
webdriver-manager # Selenium ドライバー管理
```

インストール：
```bash
pip install requests beautifulsoup4 pandas openpyxl selenium webdriver-manager
```

## 実装の特徴

### 堅牢な HTML デコーディング
- `UnicodeDammit` による自動エンコーディング検出
- netkeiba の EUC-JP ページにも対応
- フォールバック デコーディング機能

### パフォーマンス最適化
- HTTPセッション再利用（接続の使い回し）
- リトライ機能（429/500/502/503/504 エラーに対応）
- Backoff によるリトライ間隔制御

### 安全なシート名処理
- Excel のシート名制限（\\/*?:\[] など）を自動置換
- シート名の長さ制限（31文字）に対応
- シート名の衝突を自動で回避（サフィックス追加）

### Excel フォーマット
- 全セルに細い格子罫線を自動適用
- データの視認性を向上

## エラーハンドリング

- WIN5 レース取得失敗時は即座に終了（exit code 2）
- レースページ取得失敗は該当レースをスキップ（[SKIP] と表示）
- エラー情報は実行ログに出力
- テンプレートファイルが見つからない場合はエラー終了（exit code 3）

## 環境変数と定数

- `HEADERS` - User-Agent など HTTP ヘッダ
- `PC_URL` - デフォルト WIN5 ページ URL（idx=0: 土曜日、idx=1: 日曜日）
- `RACE_ID_RE` - race_id 抽出用正規表現（`race_id=\d{12}`）

## 技術仕様

- **レース ID 形式**: 12 桁数字（YYYYMMDD + 開催地コード + レース番号など）
- **日付形式**: YYYYMMDD（例：20250315）
- **タイムアウト**: HTTP 通信 15 秒
- **リトライ**: 最大 3 回（backoff_factor=0.3）

## トラブルシューティング

### "Shutuba_Past5_Table が見つかりませんでした"
- レースページ構造が変わった可能性
- netkeiba.com の仕様変更を確認

### "テンプレートが見つかりません"
- このスクリプトはテンプレートを使用していません
- 直接新規 Excel ファイルを生成します

### WIN5 race_id が取得できない
- WIN5 対象外の期間（例：休止日）
- netkeiba.com の接続障害
- URL の idx パラメータを確認

