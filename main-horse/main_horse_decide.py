# -*- coding: utf-8 -*-
import os
import re
import sys
import time
import math
import datetime as dt
import pandas as pd
import requests

from openpyxl.styles import Border, Side
from io import StringIO
from pathlib import Path
from bs4 import BeautifulSoup
from bs4 import UnicodeDammit
from urllib.parse import urlparse, parse_qs, unquote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook

# ===================== 定数 =====================
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/120.0.0.0 Safari/537.36"),
    "Referer": "https://www.netkeiba.com/",
    "Accept-Language": "ja,en;q=0.9",
}
idx = 0  # 土曜日はidx=0、日曜日はidx=1
PC_URL = f"https://race.netkeiba.com/top/win5.html?idx={idx}"
RACE_ID_RE = re.compile(r"race_id=(\d{12})")

# テンプレートファイル（スクリプトと同じフォルダに置く）
TEMPLATE_XLSX = Path(__file__).resolve().with_name("Win5軸馬決定_テンプレート.xlsx")
# ===================== 定数 =====================

# ===================== 高速化：HTTPセッション =====================
def build_session() -> requests.Session:
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry

    s = requests.Session()
    s.headers.update(HEADERS)
    retry = Retry(
        total=3, backoff_factor=0.3,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET",)
    )
    s.mount("https://", HTTPAdapter(max_retries=retry))
    s.mount("http://", HTTPAdapter(max_retries=retry))
    return s

SESSION = build_session()
# ===================== 高速化：HTTPセッション =====================

# ===================== HTMLユーティリティ =====================
SESSION = build_session()
def _decode_html_bytes(b: bytes, fallback: str = "utf-8") -> str:
    dammit = UnicodeDammit(b, is_html=True)
    if dammit.unicode_markup:
        return dammit.unicode_markup
    return b.decode(fallback, errors="replace")

def _get_html(url: str, timeout: int = 15) -> str:
    r = SESSION.get(url, timeout=timeout)
    r.raise_for_status()
    return _decode_html_bytes(r.content)
# ===================== HTMLユーティリティ =====================

# ===================== WIN5 race_idとrace_date 抽出 =====================
def _extract_ids_from_html(html: str) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    ids, seen = [], set()
    for a in soup.find_all("a", href=True):
        m = RACE_ID_RE.search(a["href"])
        if not m:
            continue
        rid = m.group(1)
        if rid not in seen:
            seen.add(rid)
            ids.append(rid)
    return ids

def _race_date(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")

    # 年は WIN5ページ中の race_id=YYYY…… から取得
    year = ""
    ids = _extract_ids_from_html(html)
    if ids:
        year = ids[0][:4]

    active_dd = soup.select_one(".RaceList_Date dl.Win5_Date dd.Active")
    if not active_dd:
        active_dd = soup.select_one(".RaceList_Date dd.Active")
    if not active_dd:
        return ""

    txt = active_dd.get_text(" ", strip=True)
    m_md = re.search(r"(\d{1,2})月\s*(\d{1,2})日", txt)
    if not m_md:
        return ""

    month = int(m_md.group(1))
    day = int(m_md.group(2))

    if not year:
        page_text = soup.get_text(" ", strip=True)
        m_y = re.search(r"(\d{4})年", page_text)
        if m_y:
            year = m_y.group(1)
        else:
            year = str(dt.datetime.now().year)

    date = f"{int(year):04d}{month:02d}{day:02d}"
    return date

def pick_win5_ids(target_url: str | None = None):
    url = target_url or PC_URL
    try:
        html = _get_html(url)
        ids = _extract_ids_from_html(html)
        date = _race_date(html)
    except Exception:
        return [], ""

    if len(ids) >= 5:
        return ids[:5], date
    return ids, date
# ===================== WIN5 race_idとrace_date 抽出 =====================

def fetch_html(url: str) -> str:
    """URLからHTML文字列を取得する"""
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    }
    resp = requests.get(url, headers=headers)
    # netkeiba は EUC-JP ヘッダだが、apparent_encoding に任せた方が楽
    resp.encoding = resp.apparent_encoding
    return resp.text


def parse_past_cell(td) -> tuple[str, str, str, str, str, str, str]:
    """
    過去走1つ分の <td class="Past"> から
    (レース名, 場所, コース, 着順, 着差, 通過順, ３F) を取り出す
    """
    if td is None:
        return "", "", "", "", "", "", ""

    # レース名（aタグ直下テキストのみ）
    race_name = ""
    a_tag = td.select_one("div.Data02 a")
    if a_tag:
        text = a_tag.find(string=True, recursive=False)
        if text:
            race_name = text.strip()

    # 日付＋場所 → 場所だけ抜く
    place = ""
    span_day_place = td.select_one("div.Data01 span:not(.Num)")
    if span_day_place:
        t = span_day_place.get_text(strip=True)
        # 例: "2025.09.15 阪神" → 最後の要素を場所とみなす
        parts = t.split()
        if len(parts) >= 2:
            place = parts[-1]

    # コース（芝1600, 芝1600(外) など）
    course = ""
    div_course = td.select_one("div.Data05")
    if div_course:
        t = div_course.get_text(" ", strip=True)
        # "芝1600 1:36.2 良" のような文字列 → 最初の要素だけ
        parts = t.split()
        if parts:
            course = parts[0]

    # 着順（Data01内の span.Num）
    finish = ""
    span_num = td.select_one("div.Data01 span.Num")
    if span_num:
        finish = span_num.get_text(strip=True)

    # 着差（Data07 内の (...)）
    margin = ""
    div_margin = td.select_one("div.Data07")
    if div_margin:
        t = div_margin.get_text(" ", strip=True)
        m = re.search(r"\(([^)]+)\)", t)   # 例: "(0.3)" → "0.3"
        if m:
            margin = m.group(1).strip()

    # Data06（通過順 + 3Fが入ってる想定）
    passing = ""
    last3f = ""
    div_06 = td.select_one("div.Data06")
    if div_06:
        t = div_06.get_text(" ", strip=True)  # 例: "4-3-4-3 (33.9) 524(+10)"

        # 通過順：先頭の "4-3-4-3" を取る（括弧の前）
        # パターンが崩れても split の先頭で拾えるようにする
        if t:
            first = t.split()[0]
            # 先頭が "4-3-4-3" 形式のときだけ採用（安全策）
            if re.fullmatch(r"\d+(?:-\d+)+", first):
                passing = first

        # 3F：括弧内数値
        m = re.search(r"\(([\d\.]+)\)", t)
        if m:
            last3f = m.group(1)

    return race_name, place, course, finish, margin, passing, last3f

# ===================== レースメタ情報抽出 =====================
def _extract_race_meta(html: str) -> tuple[str, str, str, str, str, str]:
    soup = BeautifulSoup(html, "html.parser")

    name_el = soup.select_one(".RaceName")
    name = name_el.get_text(strip=True) if name_el else ""

    d1_el = soup.select_one(".RaceData01")
    d2_el = soup.select_one(".RaceData02")
    d1 = d1_el.get_text(" ", strip=True) if d1_el else ""
    d2 = d2_el.get_text(" ", strip=True) if d2_el else ""

    year = ""
    m_id = RACE_ID_RE.search(html)
    if m_id:
        rid = m_id.group(1)
        year = rid[:4]

    race_date = ""
    active_dd = soup.select_one("#RaceList_DateList dd.Active")
    if active_dd and year:
        txt = active_dd.get_text(" ", strip=True)
        m_md = re.search(r"(\d{1,2})\s*/\s*(\d{1,2})", txt)
        if m_md:
            month = int(m_md.group(1))
            day = int(m_md.group(2))
            race_date = f"{int(year):04d}{month:02d}{day:02d}"

    if not race_date:
        date_text_candidates: list[str] = []
        date_el = soup.select_one(".RaceList_Date")
        if date_el:
            date_text_candidates.append(date_el.get_text(" ", strip=True))
        if d1:
            date_text_candidates.append(d1)
        if d2:
            date_text_candidates.append(d2)

        for txt in date_text_candidates:
            m = re.search(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日", txt)
            if m:
                y, mth, d = m.groups()
                race_date = f"{int(y):04d}{int(mth):02d}{int(d):02d}"
                break

    place = ""
    if d2:
        PLACE_PATTERN = re.compile(r"(札幌|函館|福島|新潟|東京|中山|中京|京都|阪神|小倉)")
        m = PLACE_PATTERN.search(d2)
        if m:
            place = m.group(1)

    rnum = ""
    rnum_el = soup.select_one(".RaceNum")
    rnum_text = rnum_el.get_text(strip=True) if rnum_el else ""
    m_r = re.search(r"(\d+)R", rnum_text)
    if not m_r:
        m_r = re.search(r"(\d+)R", soup.get_text(" ", strip=True))
    if m_r:
        rnum = m_r.group(1) + "R"

    return race_date, name, place, rnum

# ===================== サイトからデータ取得 =====================


def fetch_shutsuba_with_meta(url: str, timeout_sec: int = 15):
    html = _get_html(url, timeout=timeout_sec)
    race_date, name,place, rnum = _extract_race_meta(html)
    return race_date, name,place, rnum

def extract_horse_table(html: str) -> pd.DataFrame:
    """
    馬柱(5走)テーブルから
    馬番, 馬名, 性齢, 騎手名,
    前走/2走/3走/4走の(レース名, 場所, コース, 着順,3F)
    を DataFrame にして返す
    """
    soup = BeautifulSoup(html, "html.parser")

    table = soup.select_one("table.Shutuba_Past5_Table")
    if table is None:
        raise ValueError("Shutuba_Past5_Table が見つかりませんでした")

    rows = table.select("tbody tr.HorseList")

    records = []

    for tr in rows:
        # ───────── 馬番 ─────────
        uma_no = ""
        td_umaban = tr.select_one("td.Waku")
        if not td_umaban:
            # もしクラス名が違う場合の保険
            td_umaban = tr.select_one("td.Umaban")
        if td_umaban:
            uma_no = td_umaban.get_text(strip=True)

        # 馬名（Horse_Info内の Horse02 の a）
        horse_name = ""
        a_horse = tr.select_one("td.Horse_Info div.Horse02 a")
        if a_horse:
            horse_name = a_horse.get_text(strip=True)

        # 性齢（Barei）
        sex_age = ""
        span_barei = tr.select_one("td.Jockey span.Barei")
        if span_barei:
            tmp = span_barei.get_text(strip=True)
            if tmp not in ("性齢、毛色", "勝負服", "騎手"):
                sex_age = tmp

        # 騎手名
        jockey_name = ""
        a_jockey = tr.select_one('td.Jockey a[href*="/jockey/"]')
        if a_jockey:
            jockey_name = a_jockey.get_text(strip=True)

        # 過去走（前走〜5走まで入っている想定）
        past_tds = tr.select("td.Past")

        # 取りたいのは 前走, 2走, 3走, 4走 の4つ
        labels = ["前走", "2走", "3走", "4走"]
        past_data = {}
        for idx, label in enumerate(labels):
            if idx < len(past_tds):
                race_name, place, course, finish, margin, passing, last3f = parse_past_cell(past_tds[idx])
            else:
                race_name, place, course, finish, margin, passing, last3f = "", "", "", "", "", "", ""

            past_data[f"{label}_レース名"] = race_name
            past_data[f"{label}_場所"] = place
            past_data[f"{label}_コース"] = course
            past_data[f"{label}_着順"] = finish
            past_data[f"{label}_着差"] = margin
            past_data[f"{label}_通過順"] = passing
            past_data[f"{label}_３F"] = last3f

        record = {
            "馬番": uma_no,
            "馬名": horse_name,
            "性齢": sex_age,
            "騎手名": jockey_name,
        }
        record.update(past_data)
        records.append(record)

    df = pd.DataFrame(records)

    # 欲しい列の順番を明示しておく（馬番を先頭に追加）
    cols = [
        "馬番",
        "馬名",
        "性齢",
        "騎手名",
        "前走_レース名","前走_場所","前走_コース","前走_着順","前走_着差","前走_通過順","前走_３F",
        "2走_レース名","2走_場所","2走_コース","2走_着順","2走_着差","2走_通過順","2走_３F",
        "3走_レース名","3走_場所","3走_コース","3走_着順","3走_着差","3走_通過順","3走_３F",
        "4走_レース名","4走_場所","4走_コース","4走_着順","4走_着差","4走_通過順","4走_３F",
    ]
    # 存在する列だけに絞る（念のため）
    cols = [c for c in cols if c in df.columns]
    df = df[cols]
    return df

# ===================== サイトからデータ取得 =====================

# ===================== アウトプットフォルダ作成 =====================
def output_dir() -> Path:
    try:
        base = Path(__file__).resolve().parent
    except NameError:
        base = Path.cwd()
    out = base / "output"
    out.mkdir(parents=True, exist_ok=True)
    return out
# ===================== アウトプットフォルダ作成 =====================

# ===================== シート名安全化 =====================
def safe_sheet_name(name: str, used: set[str]) -> str:
    base = re.sub(r"[\\/*?:\[\]]", "_", name).strip() or "sheet"
    base = base[:31]
    cand, i = base, 2
    while cand in used:
        suf = f"_{i}"
        cand = (base[:max(0, 31 - len(suf))] + suf)[:31]
        i += 1
    used.add(cand)
    return cand
# ===================== シート名安全化 =====================

# ===================== 罫線（格子）を付与 =====================
def add_grid_border(ws):
    """ワークシート全体に細い格子罫線を付ける"""
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    max_row = ws.max_row
    max_col = ws.max_column

    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            cell.border = border
# ===================== 罫線（格子）を付与 =====================

def main():
    # オプションで WIN5ページのURL上書きも可
    url_arg = sys.argv[1] if len(sys.argv) >= 2 else None

    # WIN5 対象レースの race_id を取得
    race_ids, race_date = pick_win5_ids(url_arg)
    if not race_ids:
        print("対象の WIN5 race_id を取得できませんでした。")
        sys.exit(2)

    # 出力ファイル名
    nowstamp = dt.datetime.now().strftime("%Y%m%d%H%M%S")
    out_dir = output_dir()
    out_xlsx = out_dir / f"Win5軸馬決定_{race_date}_{nowstamp}.xlsx"
    print(f"出力開始: {out_xlsx}")

    # テンプレート読込
    if not TEMPLATE_XLSX.exists():
        print(f"テンプレートが見つかりません: {TEMPLATE_XLSX}")
        sys.exit(3)

    wb = load_workbook(TEMPLATE_XLSX)
    sheets = wb.worksheets  # 既存シート（通常5枚）

    used_sheet_names: set[str] = set()
    errors: list[str] = []
    written = 0

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        for idx_r, rid in enumerate(race_ids):
            race_url = f"https://race.netkeiba.com/race/shutuba_past.html?race_id={rid}&rf=shutuba_submenu"
            try:
                # シート名（京都9R_秋明菊賞 のような形式にする）
                html = fetch_html(race_url)
                race_date, name,place, rnum = fetch_shutsuba_with_meta(race_url)
                sheet_title = name
                if place and rnum:
                    sheet_title = f"{place}{rnum}_{name}"
                sheet_title = safe_sheet_name(sheet_title, used_sheet_names)
                print(f"[{written+1}] {sheet_title} に書き込み中…")

                df = extract_horse_table(html)
                # ファイル名は適当に。race_idを使ってもOK
                df.to_excel(writer, sheet_name=sheet_title, index=False)
                ws = writer.sheets[sheet_title]
                add_grid_border(ws)
                print(f"[{written+1}] {sheet_title} に書き込み完了")     
                written += 1
            except Exception as e:
                msg = f"{rid}: {type(e).__name__}: {e}"
                print("[SKIP]", msg)
                errors.append(msg)
    print(f"出力完了: {out_xlsx}")


if __name__ == "__main__":
    main()
