#!/usr/bin/env python3
"""
Pincode relay server.

Lokaal draaien (LAN):
    python pincode_server.py            (poort 8420)
    python pincode_server.py 9000       (andere poort)

Op een eigen server draaien (bv. bereikbaar over internet):
    PORT=8420 PINCODE_KEY=kies-hier-een-eigen-wachtwoord python3 pincode_server.py

  PORT          - poort om op te luisteren (env var wint; anders CLI-argument; anders 8420)
  PINCODE_KEY   - optioneel toegangswachtwoord. Als gezet, moet iedere URL
                  ?key=<PINCODE_KEY> bevatten, anders krijg je een 403.
                  Sterk aangeraden zodra deze server buiten je eigen LAN
                  bereikbaar is - anders kan iedereen die de URL raadt
                  ritnummers naar je scherm sturen.

Scherm:  http://<server>:<poort>/display?key=<PINCODE_KEY>
Invoer:  http://<server>:<poort>/input?key=<PINCODE_KEY>

Het Excel-bestand zelf wordt nergens door deze server gelezen - dat gebeurt
alleen lokaal in de browser van het displayscherm (via de bestandskiezer
daar). Over de server gaat alleen het opgezochte ritnummer, nooit de
pincode zelf.

Vereist alleen de Python standaardbibliotheek - er hoeft niets extra
geinstalleerd te worden (geen pip install nodig).
"""
import json
import os
import socket
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

PORT = int(os.environ.get("PORT") or (sys.argv[1] if len(sys.argv) > 1 else 8420))
ACCESS_KEY = os.environ.get("PINCODE_KEY", "").strip()

state_lock = threading.Lock()
state = {"id": 0, "query": "", "ts": 0}


def get_lan_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip


DISPLAY_HTML = r"""<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Pincode Opzoeken - Scherm</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root{
    --bg:#0f1115; --panel:#181b22; --panel2:#20242e; --border:#2a2f3a;
    --text:#e8eaef; --muted:#8b92a3; --accent:#f2a900; --good:#28c76f; --bad:#ea5455;
  }
  *{box-sizing:border-box;}
  html,body{height:100%;margin:0;background:var(--bg);color:var(--text);
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;}
  body{display:flex;flex-direction:column;min-height:100vh;}
  header{padding:16px 24px;display:flex;align-items:center;justify-content:space-between;
    border-bottom:1px solid var(--border);background:var(--panel);}
  header h1{font-size:20px;margin:0;font-weight:600;letter-spacing:.02em;}
  .file-controls{display:flex;align-items:center;gap:10px;font-size:13px;color:var(--muted);flex-wrap:wrap;}
  button{font:inherit;cursor:pointer;border:1px solid var(--border);background:var(--panel2);
    color:var(--text);padding:8px 14px;border-radius:8px;font-size:13px;transition:background .15s;}
  button:hover{background:#282d3a;}
  button.primary{background:var(--accent);color:#1a1200;border-color:var(--accent);font-weight:600;}
  button.primary:hover{filter:brightness(1.08);}
  #fileStatus,#linkStatus{white-space:nowrap;}
  main{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;
    padding:32px 16px;gap:28px;}
  .search-box{width:100%;max-width:640px;display:flex;gap:10px;}
  #ritInput{flex:1;font-size:26px;padding:16px 18px;border-radius:12px;border:2px solid var(--border);
    background:var(--panel2);color:var(--text);outline:none;text-align:center;letter-spacing:.04em;}
  #ritInput:focus{border-color:var(--accent);}
  #zoekBtn{font-size:20px;padding:16px 26px;border-radius:12px;}
  .result{width:100%;max-width:640px;background:var(--panel);border:1px solid var(--border);
    border-radius:16px;padding:36px;text-align:center;min-height:220px;display:flex;
    flex-direction:column;align-items:center;justify-content:center;gap:10px;}
  .result.empty{color:var(--muted);font-size:18px;}
  .result .label{font-size:14px;text-transform:uppercase;letter-spacing:.12em;color:var(--muted);}
  .result .pincode-value{font-size:96px;font-weight:800;letter-spacing:.06em;color:var(--good);
    line-height:1;white-space:nowrap;max-width:100%;display:inline-block;}
  .result.not-found .not-found-msg{color:var(--bad);font-size:22px;font-weight:600;}
  .candidates{width:100%;display:flex;flex-direction:column;gap:8px;text-align:left;}
  .candidates .hint{text-align:center;color:var(--muted);font-size:14px;margin-bottom:4px;}
  .candidate-btn{width:100%;display:flex;justify-content:space-between;gap:14px;padding:12px 16px;
    font-size:16px;text-align:left;}
  .candidate-btn b{color:var(--accent);}
  .candidate-btn span.c-meta{color:var(--muted);font-size:13px;}
  .meta{display:flex;flex-wrap:wrap;justify-content:center;gap:22px;margin-top:12px;font-size:15px;color:var(--muted);}
  .meta b{color:var(--text);font-weight:600;}
  footer{text-align:center;padding:12px;font-size:12px;color:var(--muted);}
  .dot{width:8px;height:8px;border-radius:50%;display:inline-block;margin-right:6px;}
  .dot.ok{background:var(--good);} .dot.no{background:var(--bad);}
</style>
</head>
<body>
<header>
  <h1>Pincode Opzoeken &mdash; Scherm</h1>
  <div class="file-controls">
    <span id="fileStatus"><span class="dot no"></span>Geen bestand geladen</span>
    <button id="loadBtn">Bestand kiezen&hellip;</button>
    <button id="refreshBtn" disabled>Ververs</button>
    <span id="linkStatus"><span class="dot no"></span>Wacht op verbinding&hellip;</span>
    <input type="file" id="fileInputFallback" accept=".xlsm,.xlsx" style="display:none">
  </div>
</header>
<main>
  <div class="search-box">
    <input type="text" id="ritInput" placeholder="Plak het A-nummer, of typ de laatste 4 cijfers&hellip;" autofocus autocomplete="off">
    <button id="zoekBtn" class="primary">Zoek</button>
  </div>
  <div class="result empty" id="resultBox">
    Laad eerst het weekbestand (.xlsm) en plak dan een ritnummer &mdash; of laat het invoerscherm het sturen.
  </div>
</main>
<footer>Leest live uit de dagbladen (Maandag t/m Zondag) van het geladen bestand &mdash; kolom C (Pincode). Ontvangt ook invoer van het invoerscherm op dit netwerk.</footer>
<script>
const ACCESS_KEY = __ACCESS_KEY_JSON__;
function withKey(url){ return ACCESS_KEY ? (url + (url.includes('?') ? '&' : '?') + 'key=' + encodeURIComponent(ACCESS_KEY)) : url; }
const DAYS = ["Maandag","Dinsdag","Woensdag","Donderdag","Vrijdag","Zaterdag","Zondag"];
let fileHandle = null;
let workbookData = null;
let autoRefreshTimer = null;
let lastSeenSubmitId = null;

const els = {
  fileStatus: document.getElementById('fileStatus'),
  loadBtn: document.getElementById('loadBtn'),
  refreshBtn: document.getElementById('refreshBtn'),
  linkStatus: document.getElementById('linkStatus'),
  fileInputFallback: document.getElementById('fileInputFallback'),
  ritInput: document.getElementById('ritInput'),
  zoekBtn: document.getElementById('zoekBtn'),
  resultBox: document.getElementById('resultBox'),
};

function idbGet(key){
  return new Promise((resolve) => {
    const req = indexedDB.open('pincode-lookup', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('handles');
    req.onsuccess = () => {
      const db = req.result;
      const tx = db.transaction('handles', 'readonly');
      const getReq = tx.objectStore('handles').get(key);
      getReq.onsuccess = () => resolve(getReq.result || null);
      getReq.onerror = () => resolve(null);
    };
    req.onerror = () => resolve(null);
  });
}
function idbSet(key, value){
  return new Promise((resolve) => {
    const req = indexedDB.open('pincode-lookup', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('handles');
    req.onsuccess = () => {
      const db = req.result;
      const tx = db.transaction('handles', 'readwrite');
      tx.objectStore('handles').put(value, key);
      tx.oncomplete = () => resolve();
    };
    req.onerror = () => resolve();
  });
}

async function tryRestoreHandle(){
  if (!window.showOpenFilePicker) return;
  const saved = await idbGet('lastFile');
  if (!saved) return;
  try{
    const perm = await saved.queryPermission({mode:'read'});
    if (perm === 'granted' || await saved.requestPermission({mode:'read'}) === 'granted'){
      fileHandle = saved;
      await loadWorkbook();
    }
  }catch(e){}
}

async function pickFile(){
  if (window.showOpenFilePicker){
    try{
      const [handle] = await window.showOpenFilePicker({
        types:[{description:'Excel werkboek', accept:{
          'application/vnd.ms-excel.sheet.macroEnabled.12':['.xlsm'],
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']
        }}],
        excludeAcceptAllOption:false, multiple:false
      });
      fileHandle = handle;
      await idbSet('lastFile', handle);
      await loadWorkbook();
    }catch(e){}
  } else {
    els.fileInputFallback.click();
  }
}

els.fileInputFallback.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const buf = await file.arrayBuffer();
  parseWorkbook(buf, file.name);
});

async function loadWorkbook(){
  if (!fileHandle) return;
  const file = await fileHandle.getFile();
  const buf = await file.arrayBuffer();
  parseWorkbook(buf, file.name);
}

function parseWorkbook(arrayBuffer, name){
  const wb = XLSX.read(arrayBuffer, {type:'array'});
  workbookData = {};
  DAYS.forEach(d => {
    const sheetName = wb.SheetNames.find(n => n.trim().toLowerCase() === d.toLowerCase());
    if (sheetName) workbookData[d] = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {header:1, raw:false, defval:''});
  });
  const now = new Date().toLocaleTimeString('nl-NL', {hour:'2-digit', minute:'2-digit', second:'2-digit'});
  els.fileStatus.innerHTML = `<span class="dot ok"></span>${name} &mdash; geladen ${now}`;
  els.refreshBtn.disabled = !fileHandle;
  startAutoRefresh();
}

function startAutoRefresh(){
  if (autoRefreshTimer) clearInterval(autoRefreshTimer);
  if (!fileHandle) return;
  autoRefreshTimer = setInterval(loadWorkbook, 60000);
}

function normalize(v){ return (v === null || v === undefined) ? '' : String(v).trim().toLowerCase(); }

function rowToRecord(dag, row){
  return { dag, ritnummer: row[0] || '', dock: row[1] || '', pincode: row[2] || '', charter: row[3] || '', naar: row[4] || '' };
}

function alleRijen(){
  const out = [];
  for (const dag of DAYS){
    const rows = workbookData[dag];
    if (!rows) continue;
    for (let r = 2; r < rows.length; r++){
      const row = rows[r] || [];
      if (normalize(row[0])) out.push(rowToRecord(dag, row));
    }
  }
  return out;
}

function zoek(){
  const query = normalize(els.ritInput.value);
  if (!query) return;
  if (!workbookData){ renderEmpty('Laad eerst het weekbestand (.xlsm).'); return; }
  const rijen = alleRijen();
  const exact = rijen.find(r => normalize(r.ritnummer) === query);
  if (exact){ renderResult(exact); els.ritInput.select(); return; }
  const suffixMatches = rijen.filter(r => { const rit = normalize(r.ritnummer); return rit && rit !== query && rit.endsWith(query); });
  if (suffixMatches.length === 1) renderResult(suffixMatches[0]);
  else if (suffixMatches.length > 1) renderCandidates(suffixMatches, query);
  else renderNotFound(query);
  els.ritInput.select();
}

function selecteerKandidaat(ritnummer){ els.ritInput.value = ritnummer; zoek(); }
window.selecteerKandidaat = selecteerKandidaat;

function renderEmpty(msg){ els.resultBox.className = 'result empty'; els.resultBox.innerHTML = msg; }

function renderNotFound(query){
  els.resultBox.className = 'result not-found';
  els.resultBox.innerHTML = `<div class="label">Niet gevonden</div>
    <div class="not-found-msg">"${escapeHtml(query.toUpperCase())}" komt niet voor in Maandag t/m Zondag</div>`;
}

function renderCandidates(matches, query){
  els.resultBox.className = 'result';
  const items = matches.map(m => `
    <button class="candidate-btn" onclick="selecteerKandidaat('${escapeHtml(String(m.ritnummer)).replace(/'/g, "&#39;")}')">
      <span><b>${escapeHtml(m.dag)}</b> &mdash; ${escapeHtml(String(m.ritnummer))}</span>
      <span class="c-meta">${escapeHtml(String(m.naar) || '—')}</span>
    </button>`).join('');
  els.resultBox.innerHTML = `<div class="candidates">
    <div class="hint">"${escapeHtml(query.toUpperCase())}" komt meerdere keren voor &mdash; kies de juiste rit:</div>
    ${items}</div>`;
}

function renderResult(found){
  els.resultBox.className = 'result';
  const pin = found.pincode !== '' ? String(found.pincode) : '—';
  els.resultBox.innerHTML = `<div class="label">Pincode</div>
    <div class="pincode-value">${escapeHtml(pin)}</div>
    <div class="meta">
      <span><b>${escapeHtml(found.dag)}</b></span>
      <span>Rit <b>${escapeHtml(String(found.ritnummer))}</b></span>
      <span>Dock <b>${escapeHtml(String(found.dock) || '—')}</b></span>
      <span>Naar <b>${escapeHtml(String(found.naar) || '—')}</b></span>
      <span>Charter <b>${escapeHtml(String(found.charter) || '—')}</b></span>
    </div>`;
  fitPincodeText();
}

function fitPincodeText(){
  const el = els.resultBox.querySelector('.pincode-value');
  if (!el) return;
  const maxWidth = els.resultBox.clientWidth - 72;
  let fontSize = 96;
  el.style.fontSize = fontSize + 'px';
  while (el.scrollWidth > maxWidth && fontSize > 24){ fontSize -= 2; el.style.fontSize = fontSize + 'px'; }
}

window.addEventListener('resize', () => { if (els.resultBox.querySelector('.pincode-value')) fitPincodeText(); });

function escapeHtml(s){ return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

els.loadBtn.addEventListener('click', pickFile);
els.refreshBtn.addEventListener('click', loadWorkbook);
els.zoekBtn.addEventListener('click', zoek);
els.ritInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') zoek(); });

// --- Polling voor invoer vanaf het andere scherm ---
async function pollLatest(){
  try{
    const res = await fetch(withKey('/api/latest'), {cache:'no-store'});
    const data = await res.json();
    els.linkStatus.innerHTML = '<span class="dot ok"></span>Verbonden';
    if (data.id && data.id !== lastSeenSubmitId){
      lastSeenSubmitId = data.id;
      if (data.query){
        els.ritInput.value = data.query;
        zoek();
      }
    }
  }catch(e){
    els.linkStatus.innerHTML = '<span class="dot no"></span>Geen verbinding met server';
  }
}
setInterval(pollLatest, 1000);
pollLatest();

tryRestoreHandle();
</script>
</body>
</html>
"""

INPUT_HTML = r"""<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Pincode Opzoeken - Invoer</title>
<style>
  :root{
    --bg:#0f1115; --panel:#181b22; --panel2:#20242e; --border:#2a2f3a;
    --text:#e8eaef; --muted:#8b92a3; --accent:#f2a900; --good:#28c76f;
  }
  *{box-sizing:border-box;}
  html,body{height:100%;margin:0;background:var(--bg);color:var(--text);
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;}
  body{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:22px;padding:24px;}
  h1{font-size:22px;margin:0 0 6px;font-weight:600;}
  #status{font-size:13px;color:var(--muted);display:flex;align-items:center;gap:6px;}
  .dot{width:8px;height:8px;border-radius:50%;display:inline-block;background:#ea5455;}
  .dot.ok{background:var(--good);}
  form{width:100%;max-width:420px;display:flex;flex-direction:column;gap:14px;}
  input{
    font-size:34px;padding:22px;border-radius:14px;border:2px solid var(--border);
    background:var(--panel2);color:var(--text);outline:none;text-align:center;letter-spacing:.05em;width:100%;
  }
  input:focus{border-color:var(--accent);}
  button{
    font-size:24px;padding:20px;border-radius:14px;border:none;background:var(--accent);
    color:#1a1200;font-weight:700;cursor:pointer;
  }
  button:active{filter:brightness(0.95);}
  #feedback{font-size:16px;min-height:24px;color:var(--good);text-align:center;}
</style>
</head>
<body>
  <h1>Pincode Invoer</h1>
  <div id="status"><span class="dot" id="statusDot"></span><span id="statusText">Verbinden&hellip;</span></div>
  <form id="form">
    <input type="text" id="q" placeholder="A-nummer of laatste 4 cijfers" autofocus autocomplete="off" inputmode="text">
    <button type="submit">Verstuur naar scherm</button>
  </form>
  <div id="feedback"></div>
<script>
const ACCESS_KEY = __ACCESS_KEY_JSON__;
function withKey(url){ return ACCESS_KEY ? (url + (url.includes('?') ? '&' : '?') + 'key=' + encodeURIComponent(ACCESS_KEY)) : url; }
const form = document.getElementById('form');
const input = document.getElementById('q');
const feedback = document.getElementById('feedback');
const statusDot = document.getElementById('statusDot');
const statusText = document.getElementById('statusText');

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const value = input.value.trim();
  if (!value) return;
  try{
    await fetch(withKey('/api/submit'), {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({query: value})
    });
    feedback.textContent = 'Verzonden: ' + value.toUpperCase();
  }catch(e){
    feedback.textContent = 'Kon niet versturen - controleer de verbinding.';
  }
  input.value = '';
  input.focus();
  setTimeout(() => { feedback.textContent = ''; }, 3000);
});

async function checkHealth(){
  try{
    await fetch(withKey('/api/health'), {cache:'no-store'});
    statusDot.classList.add('ok');
    statusText.textContent = 'Verbonden met scherm';
  }catch(e){
    statusDot.classList.remove('ok');
    statusText.textContent = 'Geen verbinding';
  }
}
setInterval(checkHealth, 4000);
checkHealth();
</script>
</body>
</html>
"""


def render_display_html():
    return DISPLAY_HTML.replace("__ACCESS_KEY_JSON__", json.dumps(ACCESS_KEY))


def render_input_html():
    return INPUT_HTML.replace("__ACCESS_KEY_JSON__", json.dumps(ACCESS_KEY))


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))

    def _send(self, code, content_type, body):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _key_ok(self, query):
        if not ACCESS_KEY:
            return True
        supplied = query.get("key", [""])[0]
        return supplied == ACCESS_KEY

    def _forbidden(self):
        self._send(
            403,
            "text/plain; charset=utf-8",
            "Toegang geweigerd: voeg ?key=<PINCODE_KEY> toe aan de URL.",
        )

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        query = parse_qs(parsed.query)

        if path in ("/", "/display", "/input", "/api/latest", "/api/health"):
            if not self._key_ok(query):
                self._forbidden()
                return

        if path in ("/", "/display"):
            self._send(200, "text/html; charset=utf-8", render_display_html())
        elif path == "/input":
            self._send(200, "text/html; charset=utf-8", render_input_html())
        elif path == "/api/latest":
            with state_lock:
                body = json.dumps(state)
            self._send(200, "application/json", body)
        elif path == "/api/health":
            self._send(200, "application/json", json.dumps({"ok": True}))
        else:
            self._send(404, "text/plain; charset=utf-8", "Niet gevonden")

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path
        query = parse_qs(parsed.query)

        if path == "/api/submit":
            if not self._key_ok(query):
                self._forbidden()
                return
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length) if length else b"{}"
            try:
                payload = json.loads(raw.decode("utf-8"))
            except Exception:
                payload = {}
            query_val = str(payload.get("query", "")).strip()
            with state_lock:
                state["id"] += 1
                state["query"] = query_val
                state["ts"] = time.time()
                new_id = state["id"]
            self._send(200, "application/json", json.dumps({"ok": True, "id": new_id}))
        else:
            self._send(404, "text/plain; charset=utf-8", "Niet gevonden")


def main():
    ip = get_lan_ip()
    suffix = f"?key={ACCESS_KEY}" if ACCESS_KEY else ""
    server = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    print("=" * 60)
    print("Pincode relay server draait.")
    print()
    print(f"  Scherm:   http://localhost:{PORT}/display{suffix}")
    print(f"  Invoer:   http://{ip}:{PORT}/input{suffix}")
    print()
    if ACCESS_KEY:
        print("Toegangssleutel (PINCODE_KEY) is actief - bovenstaande links")
        print("bevatten hem al. Deel alleen deze links, niet de sleutel apart.")
    else:
        print("LET OP: er is geen PINCODE_KEY ingesteld. Iedereen die deze")
        print("URL bereikt kan ritnummers naar het scherm sturen. Stel")
        print("PINCODE_KEY in als deze server buiten je eigen LAN bereikbaar is.")
    print()
    print("Beide apparaten moeten deze server kunnen bereiken (zelfde netwerk,")
    print("of het publieke adres van je eigen server).")
    print("Stoppen: Ctrl+C in dit venster.")
    print("=" * 60)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer gestopt.")


if __name__ == "__main__":
    main()
