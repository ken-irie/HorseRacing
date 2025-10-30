# -*- coding: utf-8 -*-
"""
WIN5ページから race_id を5つ抽出
- 既定: https://race.netkeiba.com/top/win5.html （当日）
- 任意日付: https://race.netkeiba.com/top/win5.html?date=YYYYMMDD
- モバイル版フォールバック: https://race.sp.netkeiba.com/?pid=win5&date=YYYYMMDD

使い方:
python win5_pick_ids.py
  → 当日の5レースの race_id を表示

python win5_pick_ids.py "https://race.netkeiba.com/top/win5.html?date=20251005"
  → 指定日の5レースの race_id を表示
"""
import re
import sys
import time
import datetime as dt
import requests
from bs4 import BeautifulSoup
from bs4 import UnicodeDammit

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer": "https://www.netkeiba.com/",
    "Accept-Language": "ja,en;q=0.9",
}

PC_URL = "https://race.netkeiba.com/top/win5.html"
SP_URL = "https://race.sp.netkeiba.com/?pid=win5&date={date}"  # YYYYMMDD

RACE_ID_RE = re.compile(r"race_id=(\d{12})")

def _get(url: str, timeout: int = 15) -> str:
    r = requests.get(url, headers=HEADERS, timeout=timeout)
    r.raise_for_status()
    # 文字コードを頑強に判定
    dammit = UnicodeDammit(r.content, is_html=True)
    html = dammit.unicode_markup
    if not html:
        # 最後の手段
        r.encoding = r.apparent_encoding or "utf-8"
        html = r.text
    return html

def _extract_ids_from_html(html: str) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    ids = []
    seen = set()

    # aタグのhrefから拾う（順序維持でユニーク）
    for a in soup.find_all("a", href=True):
        href = a["href"]
        m = RACE_ID_RE.search(href)
        if not m:
            continue
        rid = m.group(1)
        # 11R や 10R 以外のノイズが混ざることがあるが、WIN5ページ内では5件に収束する想定
        if rid not in seen:
            seen.add(rid)
            ids.append(rid)
    return ids

def pick_win5_ids(target_url: str | None = None) -> list[str]:
    """
    target_url 省略時は当日の PC 版URLを用いる。
    取得できなければ当日を推定して SP 版にフォールバック。
    """
    # 当日（日本時間）
    jst = dt.timezone(dt.timedelta(hours=9))
    today = dt.datetime.now(jst).strftime("%Y%m%d")

    # URL未指定ならPC版
    url = target_url or PC_URL

    # PC版を試す
    try:
        html = _get(url)
        ids = _extract_ids_from_html(html)
        if len(ids) >= 5:
            return ids[:5]
    except Exception:
        pass  # フォールバックへ

    # 日付パラメータがURLに無ければ today を付けてSP版へ
    date = today
    m = re.search(r"date=(\d{8})", url)
    if m:
        date = m.group(1)

    # モバイル版フォールバック
    try:
        sp_url = SP_URL.format(date=date)
        html = _get(sp_url)
        ids = _extract_ids_from_html(html)
        if len(ids) >= 5:
            return ids[:5]
        return ids  # 5未満でも返す（呼び出し側で扱う）
    except Exception:
        return []

def main():
    url = sys.argv[1] if len(sys.argv) >= 2 else None
    ids = pick_win5_ids(url)
    # print (ids)
    for race_id in ids:
        url = f"https://race.netkeiba.com/race/shutuba.html?race_id={race_id}"
        print(url)
if __name__ == "__main__":
    main()
