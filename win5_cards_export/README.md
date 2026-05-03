# win5_cards_export.py

netkeiba.com から WIN5 対象レースの出馬表データを取得し、既製 Excel テンプレート（`race_cards.xlsx`）の書式・計算式をそのままに自動入力するツールです。

## 概要

このスクリプトは以下の処理を自動化します：

1. **WIN5 対象レース取得** - netkeiba.com の WIN5 ページから race_id（最大 5 つ）を取得
2. **出馬表スクレイピング** - 各レースの出馬表（馬番・オッズ・馬名・性齢・斤量・騎手名）を取得
3. **テンプレート書き込み** - `race_cards.xlsx` の `オッズデータ入力` シートに直接書き込み
4. **書式・計算式の保持** - テンプレートの条件付き書式・計算式・セル結合をそのまま引き継ぎ

## 主要機能

### WIN5 情報取得
- `pick_win5_ids()` - WIN5 ページから race_id を取得（PC版 → SP版 フォールバック）
- `_extract_ids_from_html()` - HTML から race_id を抽出

### 出馬表データ取得
- `fetch_shutsuba_with_meta()` - 出馬表テーブルとレースメタ情報を取得
  - まず静的 HTML で取得、失敗時は Selenium（headless Chrome）にフォールバック
- `_extract_table()` - pandas テーブル解析（列名の自動マッピング対応）
- `_extract_race_meta()` - レース名・開催地・コース情報・発走時刻を抽出

### 抽出データ項目

| 列名 | 内容 |
|------|------|
| 人気順 | 単勝人気（RANK 計算式あり） |
| 馬番 | 馬番号 |
| オッズ | 単勝オッズ |
| 馬名 | 馬の名前 |
| 性齢 | 性別と年齢（例: 牡4） |
| 斤量 | 負担重量 |
| 騎手名 | 騎手の名前 |

### テンプレートへの書き込み
- `write_race_to_odds_sheet()` - `オッズデータ入力` シートの指定 WIN 区画にデータを書き込む
  - 結合セルは読み取り専用のためスキップ
  - `=` で始まるセル（計算式）はスキップ（テンプレートの計算式を保持）
- `_parse_race_time()` - RaceData01 から発走時刻（HH:MM）を抽出
- `_parse_course_label()` - 年齢条件・距離・芝/ダート・右/左 を組み立て

## テンプレートファイル（race_cards.xlsx）

### `オッズデータ入力` シート構成

5 レース分が**横並び**で 1 シートに格納されています。

| WIN | セクション開始列 |
|-----|--------------|
| WIN1 | B 列（2） |
| WIN2 | N 列（14） |
| WIN3 | Z 列（26） |
| WIN4 | AL 列（38） |
| WIN5 | AX 列（50） |

### 各セクションの行構成

| 行 | 内容 | 例 |
|----|------|-----|
| 4 行目 | 発走時刻 | `15:45` |
| 5 行目 | 開催地・R番・レース名 | `阪神11R_阪神大賞典` |
| 6 行目 | 年齢条件・距離・コース | `4歳以上オープン3,000メートル（芝・右）` |
| 7 行目 | ヘッダー（差/人気順/馬番/オッズ/馬名等/…） | — |
| 8〜25 行目 | データ（最大 18 頭） | — |

### 各セクション内のデータ列オフセット

| オフセット | 内容 | 備考 |
|-----------|------|------|
| +0 | 差 | **計算式（保持）** |
| +1 | 人気順 | **計算式 RANK（保持）** |
| +2 | 馬番 | ← 書き込む |
| +3 | オッズ | ← 書き込む |
| +4 | 馬名等 | ← 書き込む |
| +5 | 性齢 | ← 書き込む |
| +6 | 斤量 | ← 書き込む |
| +7 | 騎手名 | ← 書き込む |

## 使用方法

### 基本的な実行

```bash
python win5_cards_export.py
```

`idx = 1`（日曜日）の WIN5 デフォルトレースを処理します。  
土曜日にする場合はスクリプト冒頭の `idx = 0` に変更してください。

### WIN5 ページ URL を指定

```bash
python win5_cards_export.py "https://race.netkeiba.com/top/win5.html?idx=0"
```

## 入出力

### 入力
- **テンプレートファイル**（必須）
  - ファイル名: `race_cards.xlsx`
  - 場所: スクリプトと同じディレクトリ
  - 必須シート: `オッズデータ入力`

### 出力
- **ファイル名**: `output/Win5出馬表_YYYYMMDD_hhmmss.xlsx`
  - テンプレートをベースに書式・計算式をそのまま保持
  - `オッズデータ入力` シートに 5 レース分のデータを横並びで入力

## 依存パッケージ

```
requests          # HTTP 通信
beautifulsoup4    # HTML パース
chardet           # エンコーディング検出（UnicodeDammit）
pandas            # テーブルパース
lxml              # HTML パーサー（pd.read_html 高速化）
selenium          # ブラウザ自動化（オッズ取得フォールバック）
openpyxl          # Excel ファイル操作（テンプレート読み書き）
webdriver-manager # ChromeDriver 管理
```

インストール：
```bash
pip install requests beautifulsoup4 chardet pandas lxml selenium openpyxl webdriver-manager
```

## 実装の特徴

### テンプレート方式
- `load_workbook("race_cards.xlsx")` でテンプレートをそのまま読み込み
- 書式・条件付き書式・計算式・セル結合を一切壊さない
- データは列名マッピングではなく**固定オフセット**で書き込み（WIN別に +2〜+7）

### 静的 HTML → Selenium フォールバック
- 静的 HTML で出馬表が取得できた場合はそのまま使用（高速）
- JavaScript レンダリングが必要な場合のみ Selenium を起動（LazyBrowser で遅延初期化）
- Selenium インスタンスは使い回し（5 レース処理中に 1 回だけ起動）

### エラーハンドリング
- 1 レースの取得失敗は `[SKIP]` で記録して次レースへ継続
- テンプレートが見つからない場合は即終了（exit code 3）

## 定数一覧

| 定数 | 値 | 説明 |
|------|----|------|
| `idx` | `1` | 0=土曜, 1=日曜 |
| `TEMPLATE_XLSX` | `race_cards.xlsx` | テンプレートファイルパス |
| `WIN_SECTION_COLS` | `[2,14,26,38,50]` | WIN1〜5 のセクション開始列 |
| `DATA_COL_OFFSETS` | `{馬番:2, …}` | セクション内データ列オフセット |
| `DATA_START_ROW` | `8` | データ開始行 |
| `DATA_END_ROW` | `25` | データ終了行（最大 18 頭） |
| `TIME_ROW` | `4` | 発走時刻を書く行 |
| `RACE_NAME_ROW` | `5` | レース名を書く行 |
| `COURSE_ROW` | `6` | コース情報を書く行 |

## トラブルシューティング

### "テンプレートが見つかりません"
- `race_cards.xlsx` がスクリプトと同じディレクトリにあることを確認

### "[SKIP] TypeError: 'in \<string\>' requires string as left operand, not float"
- pandas の `read_html` が返すテーブルの列型が float になるケース
- `_normalize_columns` で型チェック済みのため、最新版では発生しないはず

### "対象の race_id を取得できませんでした"
- WIN5 開催日でない可能性
- `idx` の値（0=土曜 / 1=日曜）を確認

### "出馬表テーブルが見つかりません"
- netkeiba.com のページ構造変更の可能性
- Selenium の ChromeDriver バージョン不一致（`webdriver-manager` で自動解決）
