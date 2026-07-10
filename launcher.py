# -*- coding: utf-8 -*-
"""WIN5 支援ツール ランチャー（ブラウザ版）

main_horse_decide.py / win5_cards_export.py をボタン1つで実行するためのGUI。
Python標準ライブラリのみで動作し、起動すると既定ブラウザに操作画面が開く。

使い方: WIN5支援ツール起動.bat をダブルクリック、または `python launcher.py`
"""
import json
import os
import socket
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

SCRIPTS = {
    "main_horse": {
        "label": "① 軸馬決定シート作成",
        "path": BASE_DIR / "main-horse" / "main_horse_decide.py",
        "output": BASE_DIR / "main-horse" / "output",
    },
    "win5_cards": {
        "label": "② WIN5出馬表エクスポート",
        "path": BASE_DIR / "win5_cards_export" / "win5_cards_export.py",
        "output": BASE_DIR / "win5_cards_export" / "output",
    },
}

# スクリプトの終了コード別のヒント（初心者向け）
EXIT_HINTS = {
    2: "ヒント: 対象レースを取得できませんでした。対象日の選択とインターネット接続を確認してください。",
    3: "ヒント: テンプレートExcelが見つかりません。スクリプトと同じフォルダにテンプレートがあるか確認してください。",
}

PREFERRED_PORT = 8765

# 背景の水墨画（ユーザー所有画像）。無くても動作する
BG_IMAGE = BASE_DIR / "assets" / "競馬.png"
# 実行中アニメ用（マイブリッジ「走る馬」1878年・パブリックドメイン）。無くても動作する
GALLOP_GIF = BASE_DIR / "assets" / "gallop.gif"
# ヘッダーのロゴ（ユーザー所有画像）。無い場合は無地の円になる
LOGO_IMAGE = BASE_DIR / "assets" / "ロゴ.png"


# ===================== 実行状態（スレッド間で共有） =====================
class AppState:
    def __init__(self):
        self.lock = threading.Lock()
        self.log_lines: list[str] = ["実行したい処理のボタンを押してください。"]
        self.running = False
        self.last_output_dir: Path | None = None

    def append(self, line: str):
        with self.lock:
            self.log_lines.append(line.rstrip("\r\n"))

    def snapshot(self, since: int) -> dict:
        with self.lock:
            return {
                "lines": self.log_lines[since:],
                "total": len(self.log_lines),
                "running": self.running,
                "canOpen": self.last_output_dir is not None and not self.running,
            }


STATE = AppState()


def run_script_thread(key: str, idx: int):
    info = SCRIPTS[key]
    STATE.append(f"=== {info['label']} を開始します（idx={idx}）===")
    env = {**os.environ, "WIN5_IDX": str(idx),
           "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
    try:
        proc = subprocess.Popen(
            [sys.executable, "-X", "utf8", str(info["path"])],
            cwd=str(info["path"].parent),
            env=env,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        for line in proc.stdout:
            STATE.append(line)
        proc.wait()
        if proc.returncode == 0:
            STATE.append("=== ✅ 正常に完了しました ===")
        else:
            STATE.append(f"=== ❌ エラーで終了しました（コード{proc.returncode}）===")
            hint = EXIT_HINTS.get(proc.returncode)
            if hint:
                STATE.append(hint)
    except Exception as e:
        STATE.append(f"=== ❌ 実行に失敗しました: {type(e).__name__}: {e} ===")
    finally:
        with STATE.lock:
            STATE.running = False


# ===================== HTML画面 =====================
PAGE_HTML = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>WIN5 支援ツール</title>
<link rel="icon" href="/logo.png">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Shippori+Mincho:wght@600;700;800&family=Zen+Kaku+Gothic+New:wght@400;500;700&family=Oswald:wght@500;600&display=swap" rel="stylesheet">
<style>
  :root {
    --paper: #f5f0e4; --ink: #231d15; --ink-soft: rgba(35, 29, 21, .25);
    --vermillion: #b23a27; --gold: #a8842c;
    --blue: #2f6bbf; --red: #c8392f; --led: #ffc233;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; min-height: 100vh;
    font-family: "Zen Kaku Gothic New", "Yu Gothic UI", "Meiryo", sans-serif;
    color: var(--ink);
    /* 水墨画を全面に敷き、中央を和紙色で淡くして可読性を確保 */
    background:
      radial-gradient(120% 95% at 50% 42%,
        rgba(247, 242, 230, .90) 30%,
        rgba(247, 242, 230, .70) 60%,
        rgba(247, 242, 230, .34) 100%),
      url('/bg.png') center / cover no-repeat fixed,
      var(--paper);
  }
  .wrap { max-width: 760px; margin: 0 auto; padding: 30px 18px 48px;
          position: relative; z-index: 1; }

  /* ---------- ヘッダー（クレスト＋題字） ---------- */
  header { display: flex; align-items: center; gap: 18px; margin-bottom: 6px;
           animation: fadeUp .6s ease both; }
  .crest { width: 66px; height: 66px; flex: none; border-radius: 50%;
           border: 2px solid rgba(35,29,21,.45);
           background: #fdfcf9; overflow: hidden;
           display: grid; place-items: center;
           box-shadow: 0 3px 14px rgba(35,29,21,.25); }
  .crest img { width: 116%; height: 116%; object-fit: contain;
               transform: translateY(-6%); }
  h1 { font-family: "Shippori Mincho", "Yu Mincho", serif; font-weight: 800;
       font-size: 1.85rem; letter-spacing: .14em; margin: 0; color: var(--ink);
       text-shadow: 0 1px 0 rgba(255,255,255,.6); }
  .sub { font-family: Oswald, sans-serif; font-weight: 500; font-size: .68rem;
         letter-spacing: .42em; color: var(--vermillion); margin-top: 4px; }

  /* ---------- 枠番8色ストライプ ---------- */
  .waku { height: 7px; border-radius: 4px; margin: 14px 0 26px;
          animation: fadeUp .6s ease .08s both;
          box-shadow: 0 2px 8px rgba(35,29,21,.3);
          outline: 1px solid rgba(35,29,21,.2);
          background: linear-gradient(90deg,
            #f5f5f5 0 12.5%, #1a1a1a 12.5% 25%, #e53935 25% 37.5%, #1e88e5 37.5% 50%,
            #fdd835 50% 62.5%, #43a047 62.5% 75%, #fb8c00 75% 87.5%, #ec407a 87.5% 100%); }

  /* ---------- カード共通（半透明の和紙パネル） ---------- */
  .card { background: rgba(253, 250, 242, .68);
          border: 1px solid var(--ink-soft); border-radius: 12px;
          padding: 18px 22px; margin-bottom: 16px;
          box-shadow: 0 10px 28px rgba(35,29,21,.16);
          backdrop-filter: blur(5px); -webkit-backdrop-filter: blur(5px);
          animation: fadeUp .6s ease both; }
  .card:nth-of-type(2) { animation-delay: .12s; }
  .card:nth-of-type(3) { animation-delay: .2s; }
  .card:nth-of-type(4) { animation-delay: .28s; }
  .card h2 { font-family: Oswald, sans-serif; font-weight: 600; font-size: .78rem;
             letter-spacing: .3em; color: var(--vermillion); margin: 0 0 14px;
             display: flex; align-items: center; gap: 10px; }
  .card h2::after { content: ""; flex: 1; height: 1px;
                    background: linear-gradient(90deg, rgba(178,58,39,.4), transparent); }

  /* ---------- 対象日チップ ---------- */
  .chips { display: flex; gap: 12px; flex-wrap: wrap; }
  .chip { position: relative; cursor: pointer; }
  .chip input[type="radio"] { position: absolute; opacity: 0; pointer-events: none; }
  .chip span { display: inline-flex; align-items: center; gap: 8px;
               padding: 10px 20px; border-radius: 999px;
               border: 1px solid rgba(35,29,21,.4); color: #56503f;
               background: rgba(255,253,246,.6);
               font-weight: 500; font-size: .95rem;
               transition: all .18s ease; }
  .chip small { opacity: .6; font-size: .72rem; letter-spacing: .05em; }
  .chip:hover span { border-color: var(--ink); transform: translateY(-1px); }
  .chip.sat input:checked + span { background: var(--blue); border-color: var(--blue);
                                   color: #fff; box-shadow: 0 4px 14px rgba(47,107,191,.4); }
  .chip.sun input:checked + span { background: var(--red); border-color: var(--red);
                                   color: #fff; box-shadow: 0 4px 14px rgba(200,57,47,.4); }
  .chip.hol input:checked + span { background: var(--gold); border-color: var(--gold);
                                   color: #fff; box-shadow: 0 4px 14px rgba(168,132,44,.4); }
  #idxInput { width: 3.2em; padding: 2px 6px; border-radius: 6px; border: 1px solid rgba(35,29,21,.3);
              font-family: Oswald, sans-serif; font-size: .95rem; text-align: center; }
  #idxInput:disabled { background: rgba(35,29,21,.08); border-color: transparent; color: #999; }

  /* ---------- 馬券風チケットボタン ---------- */
  .btn-row { display: flex; gap: 16px; flex-wrap: wrap; }
  button.run { flex: 1; min-width: 250px; cursor: pointer; text-align: left;
               border: 1px solid rgba(35,29,21,.3); border-radius: 8px; padding: 0; overflow: hidden;
               background: linear-gradient(168deg, #fffdf6, #f1e9d4);
               color: var(--ink);
               box-shadow: 0 6px 16px rgba(35,29,21,.28);
               transition: transform .16s ease, box-shadow .16s ease, filter .16s ease;
               position: relative; }
  button.run::before { /* ミシン目 */
    content: ""; position: absolute; top: 0; bottom: 0; left: 46px; width: 0;
    border-left: 2px dashed rgba(35,29,21,.3); }
  button.run::after { /* 半券の縦書きラベル */
    content: "勝馬投票券"; position: absolute; top: 50%; left: 14px;
    transform: translateY(-50%); writing-mode: vertical-rl;
    white-space: nowrap; line-height: 1;
    font-family: "Shippori Mincho", serif; font-size: .58rem; letter-spacing: .22em;
    color: rgba(35,29,21,.5); }
  .t-inner { display: block; padding: 12px 16px 13px 60px; }
  .t-head { display: flex; justify-content: space-between; align-items: baseline;
            font-family: Oswald, sans-serif; font-size: .62rem; font-weight: 600;
            letter-spacing: .28em; color: var(--gold);
            border-bottom: 1px solid rgba(35,29,21,.18); padding-bottom: 5px; }
  .t-label { display: block; font-family: "Shippori Mincho", serif; font-weight: 700;
             font-size: 1.06rem; letter-spacing: .06em; margin-top: 9px; }
  .t-sub { display: block; font-size: .72rem; color: #6b6350; margin-top: 4px; }
  button.run:hover:not(:disabled) { transform: translateY(-3px) rotate(-.4deg);
      box-shadow: 0 12px 24px rgba(35,29,21,.34), 0 0 16px rgba(178,58,39,.18); }
  button.run:active:not(:disabled) { transform: translateY(-1px); }
  button.run:disabled { filter: grayscale(.6) brightness(.92); opacity: .7; cursor: not-allowed; }

  /* ---------- 電光掲示板ログ（硯の黒） ---------- */
  .board-head { display: flex; align-items: center; gap: 10px; }
  .live { display: none; align-items: center; gap: 6px;
          font-family: Oswald, sans-serif; font-size: .68rem; letter-spacing: .25em;
          color: var(--vermillion); }
  .live::before { content: ""; width: 8px; height: 8px; border-radius: 50%;
                  background: var(--vermillion); box-shadow: 0 0 8px var(--vermillion);
                  animation: blink 1s ease infinite; }
  .live.on { display: inline-flex; }
  #log { background:
           repeating-linear-gradient(0deg, rgba(255,255,255,.025) 0 1px, transparent 1px 3px),
           #14100b;
         color: var(--led);
         font-family: Consolas, "Courier New", monospace; font-size: .84rem; line-height: 1.55;
         border: 1px solid rgba(35,29,21,.5); border-radius: 8px;
         padding: 13px 15px; height: 250px; overflow-y: auto; white-space: pre-wrap;
         text-shadow: 0 0 6px rgba(255,194,51,.35);
         box-shadow: inset 0 0 24px rgba(0,0,0,.75); }
  #log::-webkit-scrollbar { width: 8px; }
  #log::-webkit-scrollbar-thumb { background: rgba(255,194,51,.25); border-radius: 4px; }

  /* ---------- 実行中コース（マイブリッジの馬がターフを疾走） ---------- */
  #track { display: none; position: relative; height: 60px; margin-top: 12px;
           border-radius: 8px; overflow: hidden; isolation: isolate;
           background: linear-gradient(180deg, #326041 0%, #24492f 55%, #1b3a24 100%);
           border: 1px solid rgba(35,29,21,.4);
           box-shadow: inset 0 4px 16px rgba(0,0,0,.4); }
  #track.on { display: block; animation: trackIn .35s ease both; }
  @keyframes trackIn { from { opacity: 0; transform: translateY(8px); }
                       to { opacity: 1; transform: none; } }
  /* 後方へ流れるスピードライン（芝） */
  .rail { position: absolute; inset: 0; width: 200%;
          background: repeating-linear-gradient(90deg,
            rgba(255,255,255,.10) 0 38px, transparent 38px 102px);
          animation: ground .55s linear infinite; }
  @keyframes ground { to { transform: translateX(-102px); } }
  /* 内埒の白線も後方へ流す */
  #track::before { content: ""; position: absolute; left: 0; top: 72%; height: 2px;
                   width: 200%; z-index: 1;
                   background: repeating-linear-gradient(90deg,
                     rgba(255,255,255,.55) 0 26px, transparent 26px 52px);
                   animation: rails .38s linear infinite; }
  @keyframes rails { to { transform: translateX(-52px); } }
  /* ゴール板（紅白の決勝柱） */
  #track::after { content: ""; position: absolute; right: 40px; top: 0; bottom: 0;
                  width: 12px; z-index: 1; opacity: .95;
                  background: repeating-conic-gradient(#f4f0e6 0% 25%, #b23a27 25% 50%);
                  background-size: 12px 12px;
                  box-shadow: 0 0 8px rgba(0,0,0,.4); }
  .galloper { position: absolute; left: -22%; top: 50%; transform: translateY(-50%);
              z-index: 2; animation: run 3s linear infinite;
              filter: invert(1) grayscale(1) contrast(2.4) brightness(1.05);
              mix-blend-mode: screen; }
  .galloper img { height: 48px; display: block; }
  /* 巻き上がる土煙 */
  .galloper::before { content: ""; position: absolute; right: 88%; top: 52%;
                      width: 54px; height: 20px; border-radius: 50%;
                      /* 親要素ごと色反転されるため、黒→反転後に白い土煙になる */
                      background: radial-gradient(closest-side, rgba(0,0,0,.55), transparent);
                      filter: blur(3px); transform-origin: right center;
                      animation: puff .45s ease-out infinite; }
  @keyframes puff { from { opacity: .85; transform: scale(.5) translateX(0); }
                    to { opacity: 0; transform: scale(1.5) translateX(-20px); } }
  @keyframes run { 0% { left: -22%; } 100% { left: 106%; } }

  /* ---------- フッター ---------- */
  .foot { display: flex; justify-content: space-between; margin-top: 4px;
          animation: fadeUp .6s ease .34s both; }
  .foot button { padding: 9px 22px; border-radius: 8px; cursor: pointer;
                 background: rgba(255,253,246,.6); color: var(--ink);
                 border: 1px solid rgba(35,29,21,.45);
                 font-family: "Zen Kaku Gothic New", sans-serif; font-weight: 500;
                 letter-spacing: .08em; transition: all .18s ease; }
  .foot button:hover:not(:disabled) { background: var(--ink); color: var(--paper);
                                      border-color: var(--ink); }
  .foot button:disabled { opacity: .4; cursor: not-allowed; }
  #quitBtn { border-color: rgba(178,58,39,.6); color: var(--vermillion); }
  #quitBtn:hover { background: var(--vermillion); border-color: var(--vermillion); color: #fff; }
  .credit { text-align: center; margin-top: 26px; font-size: .68rem; letter-spacing: .2em;
            color: rgba(35,29,21,.45); font-family: Oswald, sans-serif; }

  @keyframes fadeUp { from { opacity: 0; transform: translateY(14px); }
                      to { opacity: 1; transform: none; } }
  @keyframes blink { 50% { opacity: .25; } }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div class="crest">
      <img src="/logo.png" alt="" onerror="this.remove()">
    </div>
    <div>
      <h1>WIN5 支援ツール</h1>
      <div class="sub">WIN5 RACING ASSISTANT</div>
    </div>
  </header>
  <div class="waku"></div>

  <div class="card">
    <h2>対象日の選択</h2>
    <div class="chips">
      <label class="chip sat"><input type="radio" name="day" value="sat" checked>
        <span>土曜日 <small>idx=0</small></span></label>
      <label class="chip sun"><input type="radio" name="day" value="sun">
        <span>日曜日 <small>idx=1</small></span></label>
      <label class="chip hol"><input type="radio" name="day" value="holiday">
        <span>祝日 idx= <input type="number" id="idxInput" value="2" min="0" disabled></span></label>
    </div>
  </div>

  <div class="card">
    <h2>出馬受付 — 実行</h2>
    <div class="btn-row">
      <button class="run" id="btn-main_horse">
        <span class="t-inner">
          <span class="t-head"><span>JRA・WIN5</span><span>TICKET 01</span></span>
          <span class="t-label">① 軸馬決定シート作成</span>
          <span class="t-sub">出走馬の過去5走成績＋単勝オッズを取得</span>
        </span>
      </button>
      <button class="run" id="btn-win5_cards">
        <span class="t-inner">
          <span class="t-head"><span>JRA・WIN5</span><span>TICKET 02</span></span>
          <span class="t-label">② WIN5出馬表エクスポート</span>
          <span class="t-sub">対象5レースの出馬表＋オッズをExcelへ</span>
        </span>
      </button>
    </div>
  </div>

  <div class="card">
    <h2 class="board-head">実行ログ <span class="live" id="liveBadge">RUNNING</span></h2>
    <div id="log"></div>
    <div id="track">
      <div class="rail"></div>
      <div class="galloper"><img src="/gallop.gif" alt="" onerror="this.parentNode.textContent='🏇'"></div>
    </div>
  </div>

  <div class="foot">
    <button id="quitBtn">終了</button>
    <button id="openBtn" disabled>出力フォルダを開く</button>
  </div>
  <div class="credit">DATA SOURCE — NETKEIBA.COM</div>
</div>

<script>
let cursor = 0;
let quitted = false;
const log = document.getElementById("log");
const openBtn = document.getElementById("openBtn");
const idxInput = document.getElementById("idxInput");
const track = document.getElementById("track");
const liveBadge = document.getElementById("liveBadge");
const runBtns = [...document.querySelectorAll("button.run")];

document.querySelectorAll('input[name="day"]').forEach(r =>
  r.addEventListener("change", () => {
    idxInput.disabled = document.querySelector('input[name="day"]:checked').value !== "holiday";
  })
);

function resolveIdx() {
  const day = document.querySelector('input[name="day"]:checked').value;
  if (day === "sat") return 0;
  if (day === "sun") return 1;
  const n = Number(idxInput.value);
  if (!Number.isInteger(n) || n < 0) {
    alert("祝日のidxは0以上の整数で入力してください。");
    return null;
  }
  return n;
}

async function run(key) {
  const idx = resolveIdx();
  if (idx === null) return;
  runBtns.forEach(b => b.disabled = true);
  await fetch("/run", { method: "POST", headers: {"Content-Type": "application/json"},
                        body: JSON.stringify({ key, idx }) });
}

document.getElementById("btn-main_horse").onclick = () => run("main_horse");
document.getElementById("btn-win5_cards").onclick = () => run("win5_cards");
openBtn.onclick = () => fetch("/open-folder", { method: "POST" });

document.getElementById("quitBtn").onclick = async () => {
  if (!confirm("ツールを終了しますか？")) return;
  quitted = true;
  try { await fetch("/quit", { method: "POST" }); } catch (e) {}
  document.body.innerHTML = "<p style='margin:40px;font-size:1.1rem;font-family:sans-serif;color:#231d15'>終了しました。このタブは閉じてください。</p>";
};

async function poll() {
  if (quitted) return;
  try {
    const res = await fetch(`/status?since=${cursor}`);
    const st = await res.json();
    if (st.lines.length) {
      log.textContent += st.lines.join("\\n") + "\\n";
      log.scrollTop = log.scrollHeight;
      cursor = st.total;
    }
    runBtns.forEach(b => b.disabled = st.running);
    openBtn.disabled = !st.canOpen;
    track.classList.toggle("on", st.running);
    liveBadge.classList.toggle("on", st.running);
  } catch (e) { /* サーバー停止中などは無視 */ }
  setTimeout(poll, 500);
}
poll();
</script>
</body>
</html>
"""


# ===================== HTTPサーバー =====================
class Handler(BaseHTTPRequestHandler):
    server_version = "Win5Launcher/1.0"

    def log_message(self, *args):
        pass  # アクセスログは出さない

    def _send(self, code: int, body: bytes, ctype: str):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, obj, code: int = 200):
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception:
            return {}

    def do_GET(self):
        if self.path == "/" or self.path.startswith("/index"):
            self._send(200, PAGE_HTML.encode("utf-8"), "text/html; charset=utf-8")
        elif self.path == "/bg.png":
            if BG_IMAGE.exists():
                self._send(200, BG_IMAGE.read_bytes(), "image/png")
            else:
                self._send(404, b"not found", "text/plain")
        elif self.path == "/gallop.gif":
            if GALLOP_GIF.exists():
                self._send(200, GALLOP_GIF.read_bytes(), "image/gif")
            else:
                self._send(404, b"not found", "text/plain")
        elif self.path == "/logo.png":
            if LOGO_IMAGE.exists():
                self._send(200, LOGO_IMAGE.read_bytes(), "image/png")
            else:
                self._send(404, b"not found", "text/plain")
        elif self.path.startswith("/status"):
            try:
                since = int(self.path.split("since=")[1])
            except (IndexError, ValueError):
                since = 0
            self._send_json(STATE.snapshot(since))
        else:
            self._send(404, b"not found", "text/plain")

    def do_POST(self):
        if self.path == "/run":
            body = self._read_json()
            key = body.get("key")
            idx = body.get("idx")
            if key not in SCRIPTS or not isinstance(idx, int) or idx < 0:
                self._send_json({"error": "bad request"}, 400)
                return
            info = SCRIPTS[key]
            if not info["path"].exists():
                STATE.append(f"❌ スクリプトが見つかりません: {info['path']}")
                self._send_json({"error": "script not found"}, 404)
                return
            with STATE.lock:
                if STATE.running:
                    self._send_json({"error": "already running"}, 409)
                    return
                STATE.running = True
                STATE.last_output_dir = info["output"]
            threading.Thread(target=run_script_thread, args=(key, idx), daemon=True).start()
            self._send_json({"ok": True})
        elif self.path == "/open-folder":
            out = STATE.last_output_dir
            if out is not None:
                out.mkdir(parents=True, exist_ok=True)
                os.startfile(str(out))
            self._send_json({"ok": True})
        elif self.path == "/quit":
            self._send_json({"ok": True})
            threading.Thread(target=self.server.shutdown, daemon=True).start()
        else:
            self._send(404, b"not found", "text/plain")


def find_port() -> int:
    """使えるなら固定ポート、ふさがっていれば空きポートを使う"""
    with socket.socket() as s:
        try:
            s.bind(("127.0.0.1", PREFERRED_PORT))
            return PREFERRED_PORT
        except OSError:
            pass
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def main():
    port = find_port()
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}/"
    print(f"WIN5 支援ツール: {url}", flush=True)

    if os.environ.get("LAUNCHER_NO_BROWSER") != "1":
        threading.Timer(0.3, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
