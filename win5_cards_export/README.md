# win5_export.py

netkeiba.com から WIN5 馬券の対象レース馬柱データと単勝オッズをリアルタイム取得し、既製 Excel テンプレートに自動入力するツールです。オッズは Selenium + headless Chrome でレンダリングして取得します。

## 概要

このスクリプトは以下の処理を自動化します：

1. **WIN5 対象レース取得** - netkeiba.com の WIN5 ページから対象レースの race_id（5つまで）を取得
2. **馬柱データ並列取得** - 最大 5 ワーカーで並列にレース情報をスクレイピング
3. **テンプレート入力** - 既製 Excel テンプレートに馬柱データを列マッピング方式で配置
4. **計算式自動挿入** - MIN、SUM、COUNTIF など分析用の計算式を自動生成
5. **単勝オッズ取得** - Selenium で shutuba.html をレンダリングして live オッズを抽出
6. **オッズ入力** - 取得したオッズを正しい行に書き込み

## 主要機能

### WIN5 情報取得
- `pick_win5_ids()` - WIN5 ページから race_id と開催日を取得（最大 5 レース）
- `_extract_ids_from_soup()` - HTML から race_id を抽出
- `_race_date_from_soup()` - 開催日を解析

### 馬柱データ取得（並列処理）
- `fetch_horse_rows()` - 1 レースの馬柱テーブル全体をパースして辞書リストで返す
- `_parse_past_td()` - 過去 1 走分のセルから詳細情報を抽出
- `_extract_place_rnum_from_soup()` - 開催地とレース番号を取得
- ThreadPoolExecutor を使った並列フェッチ（最大 5 ワーカー）

### 抽出データ項目

**基本情報**
- 馬番（枠番）
- 馬名
- 性齢
- 騎手名

**過去走データ（前走/2走/3走/4走）**
- レース名
- 開催地
- コース
- 着順
- 着差
- 通過順
- 3F タイム

### Excel テンプレート入力

- `fill_worksheet()` - データを列マッピングに従って配置、既存オッズ列を保護しながら入力
- `_row_formulas()` - 自動計算式を生成（MIN、SUM、COUNTIF）
- `fill_odds()` - オッズマップをワークシートに書き込み

### 列マッピング（DATA_COL_MAP）

| カラム名 | 列 | 例 |
|---------|-----|-----|
| 馬番 | B(2) | 3 |
| 馬名 | D(4) | サンプル号 |
| 性齢 | E(5) | 5歳牡 |
| 騎手名 | F(6) | 山田太郎 |
| 前走_レース名 | G(7) | 〇〇賞 |
| 前走_場所 | H(8) | 東京 |
| ... | ... | ... |
| 4走_３F | AN(34) | 35.2 |

**特別な列**
- A(1): 予想（手動入力欄）
- C(3): オッズ（自動入力対象）
- AI-AO(35-41): 計算式（自動生成）
- AP(42): 脚質（手動入力欄）

### オッズ取得（Selenium）

- `fetch_odds()` - headless Chrome で shutuba.html をロード
  - CSS セレクタ `[id^="odds-1_"]` から単勝オッズを抽出
  - 馬番ごとのオッズを辞書で返す
  - タイムアウト: 30 秒、ページ読み込みエラーは無視
- `_LazyBrowser` - ブラウザインスタンスの遅延初期化と再利用
  - 自動化検出回避機能（disable-automation フラグなど）
  - User-Agent スプーフィング

### 自動生成計算式

35〜41 列に以下の計算式が自動挿入されます：

```
35: =MIN(AH,AA,T,M)      # 最小オッズ
36: =SUM(K,R,Y,AF)       # 合計着差
37-40: =COUNTIF(...)     # 着順別カウント（1着～4着）
41: =SUM(AK:AN)          # 計カウント
```

## 使用方法

### 基本的な実行

```bash
python win5_export.py
```

土曜日（WIN5_IDX=0）の WIN5 デフォルトレースを処理します。

### WIN5 ページ URL を指定

```bash
python win5_export.py "https://race.netkeiba.com/top/win5.html?idx=1"
```

日曜日（idx=1）の WIN5 を処理する場合など。

## 入出力

### 入力
- **WIN5 ページ URL**（オプション）
  - デフォルト: `https://race.netkeiba.com/top/win5.html?idx=0`（土曜日）
  - 最大 5 レース分の race_id を取得
- **テンプレートファイル**（必須）
  - ファイル名: `main_horse_decide_sheets.xlsx`
  - 場所: スクリプトと同じディレクトリ
  - 形式: 5 つのシート（WIN 1〜WIN 5）、列 A〜AO にデータエリア

### 出力
- **ファイル名**: `output/main_horse_decide_sheets_YYYYMMDD_hhmmss.xlsx`
  - テンプレートをコピーして、データ＋オッズを追加
  - 5 つのシートに対応レースの馬柱データを入力
- **シート名**: `WIN 1`、`WIN 2`、...（オプションで `WIN 1_京都9R` のように修飾）

## 依存パッケージ

```
requests          # HTTP 通信
beautifulsoup4    # HTML パース
selenium          # ブラウザ自動化（オッズ取得）
openpyxl          # Excel ファイル操作
webdriver-manager # Chromedriver 管理
concurrent        # 標準ライブラリ（ThreadPoolExecutor）
```

インストール：
```bash
pip install requests beautifulsoup4 selenium openpyxl webdriver-manager
```

## 実装の特徴

### 並列処理
- ThreadPoolExecutor で最大 5 ワーカーを用いた馬柱データ並列取得
- I/O 待機時間を大幅削減（約 5 レース × 15 秒 → 15～20 秒）
- エラー時は該当レースのみスキップ

### オッズのリアルタイム取得
- Selenium + headless Chrome によるブラウザレンダリング
- JavaScript 実行後の最終的なオッズを取得（シングルオッズ単勝）
- 自動化検出回避（`disable-blink-features=AutomationControlled` など）
- タイムアウト時の無視（レース前など取得不可時）

### テンプレート方式の利点
- 計算式や書式が既に組み込まれているため、自動入力で即座に使用可能
- 列マッピングにより柔軟な配置（テンプレート変更時は DATA_COL_MAP のみ修正）
- オッズ列（C列）は自動入力でも保護可能

### エラーハンドリング
- 馬柱取得失敗は該当レースをスキップ、他レースは継続
- オッズ取得失敗はスキップ（レース前など）
- 実行ログに全エラーを累積表示
- プロセス終了時に警告サマリーを出力

## テンプレートファイル要件

### 推奨構成
- シート: WIN 1, WIN 2, WIN 3, WIN 4, WIN 5
- データ行開始: 行 2（行 1 はヘッダ）
- データ列: A〜AO（最小限）
- 推奨: A1=予想、B1=馬番、C1=オッズ、D1=馬名、..., AP1=脚質
- 計算式プレースホルダ: AI1〜AO1 にラベル、対応セルに =NA() など

### 列の詳細
```
A: 予想（手動記入欄）
B: 馬番
C: オッズ（自動入力）
D: 馬名
E: 性齢
F: 騎手名
G-M: 前走詳細（7列）
N-T: 2走詳細（7列）
U-AA: 3走詳細（7列）
AB-AH: 4走詳細（7列）
AI-AO: 計算式エリア（7列）
AP: 脚質（手動記入欄）
```

## 環境変数と定数

- `WIN5_IDX` - 0=土曜日、1=日曜日
- `PC_URL` - デフォルト WIN5 ページ
- `DATA_COL_MAP` - 列マッピング辞書
- `TEMPLATE_XLSX` - テンプレートファイルパス
- `WDM_LOG` - Chromedriver ログレベル（0=無し）

## 技術仕様

- **並列ワーカー数**: 5（ThreadPoolExecutor）
- **オッズタイムアウト**: 30 秒
- **ページ読み込みタイムアウト**: 45 秒
- **HTTP タイムアウト**: 15 秒
- **HTTP リトライ**: 最大 3 回、backoff_factor=0.3

## トラブルシューティング

### "テンプレートが見つかりません"
- `main_horse_decide_sheets.xlsx` がスクリプトと同じディレクトリにあることを確認
- ファイル名を正確に確認（大文字小文字やスペース）

### "馬柱取得に失敗"
- netkeiba.com のページ構造が変わった可能性
- ネットワーク接続を確認
- User-Agent がブロックされていないか確認

### "オッズが取得できない"
- レース開始前（オッズ公開前）の可能性
- JavaScript の実行に時間がかかっている（タイムアウト）
- 自動化検出によるブロック（`--disable-blink-features=AutomationControlled` で対応）

### "シート名が見つかりません"
- テンプレートの 5 つのシート名が `WIN 1`～`WIN 5` であることを確認
- 大文字小文字やスペースを正確に

## パフォーマンス

- **馬柱データ取得**: 並列 5 レース同時取得で約 15～20 秒
- **オッズ取得**: Selenium ブラウザ起動 + 5 レース順次取得で約 3～5 分
- **総処理時間**: 4～6 分（ネットワーク・サーバー応答に依存）

