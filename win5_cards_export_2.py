# -*- coding: utf-8 -*-
"""
WIN5の出馬表を netkeiba の出馬表ページから取得して Excel 出力
- 使い方:
  1) race_ids に WIN5 対象の race_id を5つ指定
     例: https://race.netkeiba.com/race/shutuba.html?race_id=202505040710
         ↑の "race_id=" パラメータ値（202505040710）が必要
  2) OUTPUT_PATH を任意の保存先に変更
  3) 実行
- 出力:
  Excel: 1レース=1シート
    A1: レース名
    A2: レース詳細（発走/距離/天候/馬場 等）
    A3: 開催情報（例: 4回 東京 7日目～）
    A5～: 出馬表（枠番, 馬番, 馬名, 性齢, 斤量, 騎手, 馬体重, 調教師, 単勝, 人気, 馬ID, 騎手ID）
- 備考:
  * 単勝オッズ/人気は前日や発走直前で未掲載のことがあります（取得できない時は空欄）。
  * netkeiba 側のHTML構造変更やアクセス制限により調整が必要になる場合があります。
"""

import time
import re
import os
import math
import pandas as pd
import requests
from bs4 import BeautifulSoup

# ====== 設定 ======
OUTPUT_PATH = r"D:\work\HorseRacing\WIN5_shutuba.xlsx"  # 保存先を調整
# 例としてダミーID（実運用ではWIN5当日の5レースのrace_idに差し替えてください）
race_ids = [
    "202505040710",
    "202505040811",
    "202505040912",
    "202505041011",
    "202505041212",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer": "https://www.netkeiba.com/",
}

BASE_URL = "https://race.netkeiba.com/race/shutuba.html?race_id={race_id}"

# ====== ユーティリティ ======
def _safe_get_text(el, sep=" ", strip=True):
    if not el:
        return None
    return el.get_text(sep=sep, strip=strip)

def _to_num(x):
    if x is None or x == "":
        return None
    try:
        return float(str(x).replace(",", ""))
    except ValueError:
        return None

def _extract_id_from_href(href, key):
    """
    href のクエリやパスから ID を推測して返す（見つからなければ None）
    例: '/horse/2010101234/' -> key='horse' で '2010101234'
    """
    if not href:
        return None
    m = re.search(rf"/{key}/(\d+)/", href)
    if m:
        return m.group(1)
    # /?id= のパターンなども一応考慮
    m = re.search(rf"[?&]{key}_?id=(\d+)", href)
    if m:
        return m.group(1)
    return None

# ====== パース（レース概要） ======
def parse_race_header(soup):
    """
    レース名/詳細/開催情報を取得
    """
    race_name = _safe_get_text(soup.select_one(".RaceName"))
    # RaceData01 例: 4:50発走 / ダ1600m (左) / 天候:曇 / 馬場:良
    race_data01 = soup.select_one(".RaceData01")
    race_data01_text = None
    if race_data01:
        race_data01_text = re.sub(r"\s+", " ", race_data01.get_text(" ", strip=True))

    # RaceData02 例: 4回 東京 7日目 サラ系３歳以上 ２勝クラス など
    race_data02 = soup.select_one(".RaceData02")
    race_data02_text = None
    if race_data02:
        race_data02_text = re.sub(r"\s+", " ", race_data02.get_text(" ", strip=True))

    return race_name, race_data01_text, race_data02_text

# ====== パース（出馬表本体） ======
def parse_shutuba_table(soup):
    """
    出馬表の各馬データを抽出して DataFrame を返す
    代表的な列:
      枠番, 馬番, 馬名, 性齢, 斤量, 騎手, 馬体重, 調教師, 単勝, 人気, 馬ID, 騎手ID
    ※ netkeiba のマークアップに依存、必要に応じてCSSを調整してください
    """
    rows = []
    # よくあるテーブル: table.Shutuba_Table → tbody → tr
    # ただしマークアップが変わることがあるため、fallbackも準備
    tr_list = soup.select("table.Shutuba_Table tbody tr")
    if not tr_list:
        tr_list = soup.select("tbody tr")  # ざっくりfallback

    for tr in tr_list:
        # 出走取り消しなどの行はスキップ
        if "cancel" in tr.get("class", []):
            continue

        # 枠番
        wakuban = _safe_get_text(tr.select_one("th.Waku"))
        if wakuban is None:
            wakuban = _safe_get_text(tr.select_one("th"))  # fallback

        # 馬番
        umaban = _safe_get_text(tr.select_one("td.Umaban"))
        if umaban is None:
            # たまにth側にある場合も
            tds = tr.find_all(["td", "th"], limit=3)
            if len(tds) >= 2:
                umaban = _safe_get_text(tds[1])

        # 馬名と馬ID
        horse_a = tr.select_one("span.HorseName a") or tr.select_one("td a[href*='/horse/']")
        horse_name = _safe_get_text(horse_a)
        horse_id = _extract_id_from_href(horse_a.get("href") if horse_a else None, "horse")

        # 性齢（例: 牡5、牝4）
        sexe = _safe_get_text(tr.select_one(".Barei") or tr.select_one("td:has(.Barei)"))
        if not sexe:
            # 近辺のセルから推測
            sex_age_cell = tr.find("td", string=re.compile(r"[牡牝騸]\s*\d"))
            sexe = _safe_get_text(sex_age_cell)

        # 斤量
        kinryo = _safe_get_text(tr.select_one(".Weight") or tr.select_one("td:has(.Weight)"))
        if kinryo:
            kinryo = re.sub(r"[^\d\.]", "", kinryo)  # 数字抽出

        # 騎手と騎手ID
        j_a = tr.select_one("td a[href*='/jockey/']")
        jockey = _safe_get_text(j_a)
        jockey_id = _extract_id_from_href(j_a.get("href") if j_a else None, "jockey")

        # 馬体重（例: 482(+6)）
        bweight = _safe_get_text(tr.select_one(".Weight_kg") or tr.select_one("td:has(.Weight_kg)"))
        if not bweight:
            # "馬体重" テキストが含まれるセルからも拾う
            cand = tr.find("td", string=re.compile(r"\d+\s*\(\s*[+-]\d+\s*\)"))
            bweight = _safe_get_text(cand)

        # 調教師
        trainer = _safe_get_text(tr.select_one("td a[href*='/trainer/']"))
        # 単勝/人気（前日や未確定だと空のことあり）
        odds = _safe_get_text(tr.select_one(".Odds"))
        if odds:
            odds = odds.replace(",", "")
        ninki = _safe_get_text(tr.select_one(".Ninki"))

        # 行が実質空ならスキップ
        if not any([wakuban, umaban, horse_name]):
            continue

        rows.append({
            "枠番": wakuban,
            "馬番": umaban,
            "馬名": horse_name,
            "性齢": sexe,
            "斤量": kinryo,
            "騎手": jockey,
            "馬体重": bweight,
            "調教師": trainer,
            "単勝": odds,
            "人気": ninki,
            "馬ID": horse_id,
            "騎手ID": jockey_id,
        })

    df = pd.DataFrame(rows)
    # ソート（馬番が数字っぽければ数値で）
    def _as_int(x):
        try:
            return int(str(x).strip())
        except Exception:
            return math.inf
    if not df.empty and "馬番" in df.columns:
        df["__sort_umaban"] = df["馬番"].map(_as_int)
        df = df.sort_values("__sort_umaban", kind="stable").drop(columns="__sort_umaban")
    return df.reset_index(drop=True)

# ====== メイン処理 ======
def fetch_shutuba_by_race_id(race_id, sleep_sec=1.2, timeout=15):
    """
    race_id から出馬表ページを取得して (レース概要, 出馬表DF) を返す
    """
    url = BASE_URL.format(race_id=race_id)
    r = requests.get(url, headers=HEADERS, timeout=timeout)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    race_name, race_data01, race_data02 = parse_race_header(soup)
    df = parse_shutuba_table(soup)

    # 丁寧に少しウェイト（アクセスマナー）
    time.sleep(sleep_sec)
    return {
        "race_id": race_id,
        "race_name": race_name,
        "race_data01": race_data01,
        "race_data02": race_data02,
        "df": df
    }

def export_win5_to_excel(race_ids, output_path):
    """
    race_ids（5件想定）を順に取得→Excel 1ファイルに出力
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_rows = []
        for idx, rid in enumerate(race_ids, start=1):
            info = fetch_shutuba_by_race_id(rid)
            race_name = info["race_name"] or f"RACE_{rid}"
            sheet_name = f"{idx:01d}_{race_name[:20]}"  # シート名は31文字制限に注意

            # レース概要を上部に書き、その下に出馬表を貼る
            meta = pd.DataFrame({
                "項目": ["レース名", "レース詳細", "開催情報", "race_id"],
                "値":   [info["race_name"], info["race_data01"], info["race_data02"], rid]
            })
            # メタ情報
            meta.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)
            # 出馬表（ヘッダは行5から）
            start_row = len(meta) + 2  # 1行空ける
            df = info["df"]
            if df.empty:
                # 空でもヘッダだけ
                df = pd.DataFrame(columns=["枠番","馬番","馬名","性齢","斤量","騎手","馬体重","調教師","単勝","人気","馬ID","騎手ID"])
            df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=start_row)

            # サマリー用
            summary_rows.append({
                "No": idx,
                "race_id": rid,
                "レース名": info["race_name"],
                "レース詳細": info["race_data01"],
                "開催情報": info["race_data02"],
                "出走頭数": len(df)
            })

        # まとめシート
        pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name="WIN5_まとめ")

    return output_path

if __name__ == "__main__":
    path = export_win5_to_excel(race_ids, OUTPUT_PATH)
    print(f"出力完了: {path}")
