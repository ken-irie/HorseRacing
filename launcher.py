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
<title>WIN5 支援ツール</title>
<style>
  :root { --accent: #2e7d32; --accent2: #1565c0; }
  * { box-sizing: border-box; }
  body { font-family: "Yu Gothic UI", "Meiryo", sans-serif; margin: 0;
         background: #f0f2f5; color: #333; }
  .wrap { max-width: 720px; margin: 24px auto; padding: 0 16px; }
  h1 { font-size: 1.3rem; margin: 0 0 16px; }
  .card { background: #fff; border-radius: 10px; padding: 16px 20px;
          margin-bottom: 14px; box-shadow: 0 1px 4px rgba(0,0,0,.08); }
  .card h2 { font-size: .95rem; margin: 0 0 10px; color: #666; }
  label.radio { margin-right: 18px; cursor: pointer; }
  #idxInput { width: 3em; padding: 2px 6px; }
  #idxInput:disabled { background: #eee; }
  .btn-row { display: flex; gap: 12px; }
  button.run { flex: 1; padding: 14px 8px; font-size: 1rem; border: 0;
               border-radius: 8px; color: #fff; cursor: pointer; }
  button.run:disabled { opacity: .45; cursor: not-allowed; }
  #btn-main_horse { background: var(--accent); }
  #btn-win5_cards { background: var(--accent2); }
  #log { background: #111; color: #d5f5d5; font-family: Consolas, monospace;
         font-size: .85rem; border-radius: 8px; padding: 12px;
         height: 260px; overflow-y: auto; white-space: pre-wrap; }
  .foot { display: flex; justify-content: space-between; }
  .foot button { padding: 8px 18px; border-radius: 6px; border: 1px solid #bbb;
                 background: #fff; cursor: pointer; }
  .foot button:disabled { opacity: .45; cursor: not-allowed; }
  #quitBtn { border-color: #d99; color: #b33; }
</style>
</head>
<body>
<div class="wrap">
  <h1>🏇 WIN5 支援ツール</h1>

  <div class="card">
    <h2>対象日の選択</h2>
    <label class="radio"><input type="radio" name="day" value="sat" checked> 土曜日 (idx=0)</label>
    <label class="radio"><input type="radio" name="day" value="sun"> 日曜日 (idx=1)</label>
    <label class="radio"><input type="radio" name="day" value="holiday"> 祝日 idx=
      <input type="number" id="idxInput" value="2" min="0" disabled></label>
  </div>

  <div class="card">
    <h2>実行</h2>
    <div class="btn-row">
      <button class="run" id="btn-main_horse">① 軸馬決定シート作成</button>
      <button class="run" id="btn-win5_cards">② WIN5出馬表エクスポート</button>
    </div>
  </div>

  <div class="card">
    <h2>実行ログ</h2>
    <div id="log"></div>
  </div>

  <div class="foot">
    <button id="quitBtn">終了</button>
    <button id="openBtn" disabled>出力フォルダを開く</button>
  </div>
</div>

<script>
let cursor = 0;
let quitted = false;
const log = document.getElementById("log");
const openBtn = document.getElementById("openBtn");
const idxInput = document.getElementById("idxInput");
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
  document.body.innerHTML = "<p style='margin:40px;font-size:1.1rem'>終了しました。このタブは閉じてください。</p>";
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
