# -*- coding: utf-8 -*-
import os
import re
import sys
import time
import math
import datetime as dt
import pandas as pd
import requests

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

# ===================== 定数 =====================
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/120.0.0.0 Safari/537.36"),
    "Referer": "https://www.netkeiba.com/",
    "Accept-Language": "ja,en;q=0.9",
}
idx = 1 #土曜日はidx=0、日曜日はidx=1
PC_URL = f"https://race.netkeiba.com/top/win5.html?idx={idx}"
SP_URL = "https://race.sp.netkeiba.com/?pid=win5&date={date}"  # YYYYMMDD
RACE_ID_RE = re.compile(r"race_id=(\d{12})")

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

# ===================== HTMLユーティリティ =====================
def _decode_html_bytes(b: bytes, fallback: str = "utf-8") -> str:
    dammit = UnicodeDammit(b, is_html=True)
    if dammit.unicode_markup:
        return dammit.unicode_markup
    return b.decode(fallback, errors="replace")

def _get_html(url: str, timeout: int = 15) -> str:
    r = SESSION.get(url, timeout=timeout)
    r.raise_for_status()
    return _decode_html_bytes(r.content)

# ===================== Selenium（必要時のみ） =====================
class LazyBrowser:
    """必要な時だけ起動し、プロセスは使い回す。"""
    def __init__(self):
        self._driver = None

    def _new_driver(self):
        os.environ["WDM_LOG"] = "0"
        os.environ["WDM_PRINT_FIRST_LINE"] = "False"
        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-software-rasterizer")
        options.add_argument("--disable-webgl")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")
        options.add_argument("--blink-settings=imagesEnabled=false")
        options.add_argument("--lang=ja-JP")
        options.add_argument("--remote-debugging-pipe")  # DevToolsログ抑制
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        options.add_argument("--log-level=3")
        options.add_argument("--silent")
        service = ChromeService(
            ChromeDriverManager().install(),
            log_output=open(os.devnull, "w", encoding="utf-8", errors="ignore")
        )
        d = webdriver.Chrome(service=service, options=options)
        d.set_page_load_timeout(45)
        d.set_script_timeout(45)
        return d

    @property
    def driver(self):
        if self._driver is None:
            self._driver = self._new_driver()
        return self._driver

    def get_rendered_html(self, url: str, wait_css: str = None, hard_timeout: int = 25, wait_odds: bool = False) -> str:
        d = self.driver
        try:
            try:
                d.get(url)
            except TimeoutException:
                pass

            t0 = time.time()
            while time.time() - t0 < min(8, hard_timeout):  # DOMContentLoaded 相当を短めに
                if d.execute_script("return document.readyState") in ("interactive", "complete"):
                    break
                time.sleep(0.2)

            if wait_css:
                WebDriverWait(d, hard_timeout).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, wait_css))
                )

            if wait_odds:
                def odds_ready(drv):
                    try:
                        return drv.execute_script("""
                            const nodes = document.querySelectorAll('td.Popular, td.Odds, .Popular, .Odds');
                            for (const n of nodes) {
                              const t=(n.textContent||'').trim();
                              if (/^\\d+(?:\\.\\d+)?(?:\\s*倍)?$/.test(t)) return true;
                            }
                            return false;
                        """)
                    except Exception:
                        return False
                WebDriverWait(d, hard_timeout).until(lambda drv: odds_ready(drv))

            if d.execute_script("return document.readyState") != "complete":
                d.execute_script("window.stop();")
            return d.page_source
        except Exception:
            return d.page_source
    def close(self):
        try:
            if self._driver:
                self._driver.quit()
        except Exception:
            pass

BROWSER = LazyBrowser()

# ===================== パース =====================
def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        cols = []
        for t in df.columns:
            parts = [str(x).strip() for x in t if str(x).strip()]
            s = " ".join(parts) if parts else ""
            toks = s.split()
            if len(toks) >= 2 and len(set(toks)) == 1:
                s = toks[0]
            cols.append(s)
        df.columns = cols
    else:
        df.columns = [str(c).strip() for c in df.columns]

    seen, uniq = {}, []
    for c in df.columns:
        if c in seen:
            seen[c] += 1
            uniq.append(f"{c}.{seen[c]}")
        else:
            seen[c] = 0
            uniq.append(c)
    df.columns = uniq

    if len(df) > 0:
        row0 = df.iloc[0].astype(str).str.replace(r"\s+", "", regex=True).tolist()
        col0 = [c.replace(" ", "") for c in df.columns]
        match_cnt = sum(1 for x in row0 if any(x and x in c for c in col0))
        if match_cnt >= max(2, len(col0)//2):
            df = df.iloc[1:].reset_index(drop=True)
    return df

def _extract_table(html: str) -> pd.DataFrame | None:
    # pandas はファイルライクの方が速い
    bio = StringIO(html)
    try:
        tables = pd.read_html(bio, flavor="lxml")
    except Exception:
        bio.seek(0)
        tables = pd.read_html(bio)

    REQUIRED = {"馬番", "人気順", "オッズ", "馬名", "騎手名", "斤量", "性齢"}
    
    col_patterns = {
        "馬番":   re.compile(r"(馬\s*番|枠\s*番|馬番|枠番|\b馬\s*#?)", re.I),
        "人気順": re.compile(r"(人気|単勝人気)", re.I),
        "オッズ": re.compile(r"(オッズ|単勝)", re.I),
        "馬名":   re.compile(r"(馬\s*名|馬名|名前)", re.I),
        "騎手名": re.compile(r"(騎手|騎手名|ジョッキー)", re.I),
        "斤量":   re.compile(r"(斤量|負担重量|負担重|重量)", re.I),
        "性齢":   re.compile(r"(性\s*齢|性齢|性別?\s*年齢|年齢\s*[／/]\s*性別?)", re.I),
    }

    def pick(df: pd.DataFrame) -> pd.DataFrame | None:
        df = _normalize_columns(df)
        cols = [str(c) for c in df.columns]
        mapping = {}
        for want, pat in col_patterns.items():
            hit = next((c for c in cols if pat.search(c)), None)
            if hit:
                mapping[hit] = want

        # 性と年齢が別カラムの表に対するフォールバック
        if "性齢" not in mapping.values():
            sex_col = next((c for c in cols if re.fullmatch(r"(性|性別)", c, re.I)), None)
            age_col = next((c for c in cols if re.fullmatch(r"(年齢|年令|age)", c, re.I)), None)
            if sex_col and age_col:
                # 一時列を作って性齢として扱う（例: 牡 + 3 → 牡3）
                df["_tmp_性齢"] = (
                    df[sex_col].astype(str).str.extract(r"(牡|牝|セ|騸|騙)", expand=False).fillna("")
                    + df[age_col].astype(str).str.extract(r"(\d+)", expand=False).fillna("")
                )
                mapping["_tmp_性齢"] = "性齢"

        # 足りない時だけ補完（人気/オッズ/騎手/斤量の推定）
        if len(set(mapping.values())) < len(REQUIRED):
            for c in cols:
                if re.search(r"(印|予想印)", c) and "人気順" not in mapping.values():
                    mapping[c] = "人気順"
                if re.search(r"(単勝|勝率|オッズ)", c, re.I) and "オッズ" not in mapping.values():
                    mapping[c] = "オッズ"
                if re.search(r"(騎手|ジョッキー)", c) and "騎手名" not in mapping.values():
                    mapping[c] = "騎手名"
                if re.search(r"(斤量|負担重量|負担重|重量)", c) and "斤量" not in mapping.values():
                    mapping[c] = "斤量"

        # すべて揃ったら正規化して返す
        if REQUIRED.issubset(set(mapping.values())):
            out = df[list(mapping.keys())].rename(columns=mapping).copy()

            # ベクトル化正規化
            out["人気順"] = pd.to_numeric(
                out["人気順"].astype(str).str.extract(r"(\d+)", expand=False),
                errors="coerce"
            )
            out["オッズ"] = pd.to_numeric(
                (out["オッズ"].astype(str)
                    .str.replace("倍", "", regex=False)
                    .str.replace(",", "", regex=False)),
                errors="coerce"
            )
            out["馬番"] = (out["馬番"].astype(str)
                .str.extract(r"(\d+)", expand=False).astype("Int64"))
            out["騎手名"] = (out["騎手名"].astype(str)
                .str.replace(r"\s+", " ", regex=True).str.strip())
            out["斤量"] = pd.to_numeric(
                out["斤量"].astype(str).str.extract(r"(\d+(?:\.\d+)?)", expand=False),
                errors="coerce"
            )
            out["馬名"] = out["馬名"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()

            # 性齢の正規化と分解（おまけ）
            out["性齢"] = out["性齢"].astype(str).str.replace(r"\s+", "", regex=True)

            # 見やすい並びにして返す（必要に応じて変更OK）
            order = [c for c in [ "人気順", "馬番", "オッズ", "馬名", "性齢", "斤量", "騎手名"] if c in out.columns]
            out = out[order]

            return out
        return None

    for tb in tables:
        got = pick(tb)
        if got is not None:
            return got
    return None

# def _extract_race_meta(html: str) -> tuple[str|None, str|None, str|None]:
#     soup = BeautifulSoup(html, "lxml")
#     name = soup.select_one(".RaceName")
#     data01 = soup.select_one(".RaceData01")
#     data02 = soup.select_one(".RaceData02")
#     name = name.get_text(strip=True) if name else None
#     data01 = re.sub(r"\s+", " ", data01.get_text(" ", strip=True)) if data01 else None
#     data02 = re.sub(r"\s+", " ", data02.get_text(" ", strip=True)) if data02 else None
#     return name, data01, data02

def _extract_race_meta(html: str) -> tuple[str|None, str|None, str|None, str|None, str|None, str|None]:
    soup = BeautifulSoup(html, "lxml")
    name = soup.select_one(".RaceName")
    data01 = soup.select_one(".RaceData01")
    data02 = soup.select_one(".RaceData02")
    rnum  = soup.select_one(".RaceNum")     # 例: 10R

    # テキスト化
    name  = name.get_text(strip=True) if name else None
    data01 = re.sub(r"\s+", " ", data01.get_text(" ", strip=True)) if data01 else None
    data02 = re.sub(r"\s+", " ", data02.get_text(" ", strip=True)) if data02 else None
    rnum  = rnum.get_text(strip=True) if rnum else None
    place: str | None = None
    PLACE_PATTERN = re.compile(r"(札幌|函館|福島|新潟|東京|中山|中京|京都|阪神|小倉)")

    # rnum 正規化（"第10R" → "10R" など）
    if rnum:
        m = re.search(r"(\d+)\s*R", rnum, flags=re.I)
        if m:
            rnum = f"{int(m.group(1))}R"

    if data02:
        m = PLACE_PATTERN.search(data02)
        if m:
            place = m.group(1)

    # --- race_date 抽出 (yyyymmdd) ---
    race_date: str | None = None

    # 1) 画面上の日本語日付から取得（例: "2025年5月5日"）
    date_hints_selectors = [
        ".RaceList_Date",   # まずここを試す
        ".RaceData01",      # ここに含まれているケースもある
        ".RaceData02",
    ]
    for sel in date_hints_selectors:
        node = soup.select_one(sel)
        if not node:
            continue
        txt = node.get_text(" ", strip=True)
        m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", txt)
        if m:
            y, mo, d = map(int, m.groups())
            race_date = f"{y:04d}{mo:02d}{d:02d}"
            break

    # 2) ダメなら <script> 内の埋め込み値から拾う（例: "kaisai_date":"20250505"）
    if not race_date:
        for s in soup.find_all("script"):
            st = s.get_text(" ", strip=True)
            m = re.search(r'"kaisai_date"\s*:\s*"(\d{8})"', st)
            if not m:
                m = re.search(r'kaisaiDate\s*[:=]\s*"(\d{8})"', st)
            if m:
                race_date = m.group(1)
                break

    return race_date, name, data01, data02, place, rnum

def fetch_shutsuba_with_meta(url: str, timeout_sec: int = 15) -> tuple[pd.DataFrame, tuple[str,str,str]]:
    # まず静的HTML
    html = _get_html(url, timeout=timeout_sec)
    df = _extract_table(html)
    race_date, name, d1, d2, place, rnum = _extract_race_meta(html)
    if df is not None and name and d1 and d2:
        return df, (race_date, name, d1, d2, place, rnum)

    # ダメなら Selenium（1インスタンス使い回し）
    html2 = BROWSER.get_rendered_html(
        url,
        wait_css=".Shutuba_Table, table.RaceTable01, .RaceTable01",
        hard_timeout=30,
        wait_odds=True
    )
    df2 = _extract_table(html2)
    race_date2, name2, d12, d22, place2, rnum2  = _extract_race_meta(html2)
    if df2 is not None and name2 and d12 and d22:
        return df2, (race_date2, name2, d12, d22, place2, rnum2)

    raise ValueError("出馬表テーブルが見つかりません。")

# ===================== WIN5 race_id 抽出（PC→SP フォールバック） =====================
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

def pick_win5_ids(target_url: str | None = None) -> list[str]:
    jst = dt.timezone(dt.timedelta(hours=9))
    today = dt.datetime.now(jst).strftime("%Y%m%d")
    url = target_url or PC_URL

    try:
        ids = _extract_ids_from_html(_get_html(url))
        if len(ids) >= 5:
            return ids[:5]
    except Exception:
        pass

    date = re.search(r"date=(\d{8})", url)
    date = date.group(1) if date else today

    try:
        ids = _extract_ids_from_html(_get_html(SP_URL.format(date=date)))
        return ids[:5] if len(ids) >= 5 else ids
    except Exception:
        return []

# ===================== Excel 出力（高速：xlsxwriter 利用） =====================
def write_sheet_as_table(writer, df: pd.DataFrame, sheet_name: str):
    # pandas書き込み → xlsxwriter でテーブル化＆書式付け
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    wb  = writer.book
    ws  = writer.sheets[sheet_name]

    nrows, ncols = df.shape
    # 見出しとバンドを xlsxwriter の Table 機能で
    ws.add_table(0, 0, nrows, ncols-1, {
        "columns": [{"header": c} for c in df.columns],
        "style": "Table Style Medium 9"  # 罫線＋交互色。高速。
    })

    # 列幅（おおよその最適化）
    for j, col in enumerate(df.columns):
        width = max(8, min(28, int(df[col].astype(str).map(len).quantile(0.8)) + 2))
        ws.set_column(j, j, width)

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

def get_output_dir() -> Path:
    try:
        base = Path(__file__).resolve().parent
    except NameError:
        base = Path.cwd()
    out = base / "output"
    out.mkdir(parents=True, exist_ok=True)
    return out

# ===================== メイン =====================
def main():
    # race_id 抽出
    url_arg = sys.argv[1] if len(sys.argv) >= 2 else None
    race_ids = pick_win5_ids(url_arg)
    if not race_ids:
        print("対象の race_id を取得できませんでした。")
        sys.exit(2)

    race_date = dt.date.today().strftime("%Y%m%d")
    nowstamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    outdir = get_output_dir()
    out_xlsx = outdir / f"Win5出馬表_{nowstamp}.xlsx"
    print(f"出力開始: {out_xlsx}")

    used_sheet_names: set[str] = set()
    errors = []
    written = 0

    # xlsxwriter を使う（装飾を後から openpyxl で全セル処理しない）
    with pd.ExcelWriter(out_xlsx, engine="xlsxwriter") as writer:
        for rid in race_ids:
            url = f"https://race.netkeiba.com/race/shutuba.html?race_id={rid}"
            try:
                df, meta = fetch_shutsuba_with_meta(url)
                race_date, name, d1, d2, place, rnum = meta
                if not (name and d1 and d2):
                    raise ValueError("race meta not found")
                
                # ▼シート名を「開催場所 + R番 _ レース名」にする（場所・R番が無い時はレース名のみ）
                sheet_title = name
                if place and rnum:
                    sheet_title = f"{place}{rnum}_{name}"
                sheet = safe_sheet_name(sheet_title, used_sheet_names)
                print(f"第{written+1}レース [{sheet}] シートに書き込み中…")

                # 並べ替え（馬番安定）
                keys = [c for c in ["人気順", "馬番"] if c in df.columns]
                if keys:
                    df = df.sort_values(keys, na_position="last", ignore_index=True, kind="mergesort")
                # df = df.sort_values(["人気順", "馬番"], na_position="last", ignore_index=True)

                write_sheet_as_table(writer, df, sheet)
                print(f"第{written+1}レース [{sheet}] シートに書き込み完了")
                written += 1
            except Exception as e:
                msg = f"{rid}: {type(e).__name__}: {e}"
                print("[SKIP]", msg)
                errors.append(msg)

        if written == 0:
            pd.DataFrame({"info": ["no sheets written"]}).to_excel(writer, index=False, sheet_name="empty")
        if errors:
            pd.DataFrame({"errors": errors}).to_excel(writer, index=False, sheet_name="log")

    # 使ったら閉じる（Selenium 起動していれば）
    BROWSER.close()
    print(f"出力完了: {out_xlsx}")

if __name__ == "__main__":
    main()