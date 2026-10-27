# -*- coding: utf-8 -*-
"""
Рассылочный помощник + фоновый сендер в Telegram (Telethon).
Ключи читаются только из переменных окружения (Render → Environment).
Логин в Telegram делается через /setup?token=... (VPN не нужен).
"""
import csv
import io
import os
import re
import secrets
import socket
import threading
import time
import asyncio
import webbrowser
from urllib.parse import urlparse

from flask import Flask, jsonify, render_template_string, request

try:
    from telethon import TelegramClient, errors
    from telethon.sessions import StringSession
    from telethon.tl.types import User
    TELETHON_OK = True
except Exception:
    TELETHON_OK = False

app = Flask(__name__)

# ====================== НАСТРОЙКИ ИЗ ОКРУЖЕНИЯ ======================
TG_API_ID = os.environ.get("TG_API_ID")
TG_API_HASH = os.environ.get("TG_API_HASH")
SEND_DELAY = int(os.environ.get("SEND_DELAY", "60"))

# Токен для доступа к /setup. Если не задан — генерируется и печатается в логи.
SETUP_TOKEN = os.environ.get("SETUP_TOKEN") or secrets.token_urlsafe(18)
print(f"[setup] SETUP_TOKEN = {SETUP_TOKEN}", flush=True)

# Строка сессии: сначала из env, потом из памяти (после /setup без рестарта)
_runtime_session_string = os.environ.get("TG_SESSION_STRING", "") or ""

# ====================== ФИЛЬТРЫ ======================
NOT_A_SITE = {
    "vk.com", "vk.ru", "m.vk.com", "instagram.com", "facebook.com", "fb.com", "ok.ru",
    "t.me", "telegram.me", "telegram.dog", "wa.me", "whatsapp.com", "api.whatsapp.com",
    "max.ru", "youtube.com", "youtu.be", "tiktok.com", "twitter.com", "x.com",
    "taplink.cc", "taplink.ru", "linktr.ee", "yandex.ru", "ya.ru", "2gis.ru",
    "zoon.ru", "profi.ru", "avito.ru", "dikidi.net", "dikidi.ru", "yclients.com",
}

TG_SKIP = {"joinchat", "s", "addstickers", "share", "c", "proxy", "socks", "login", "iv", "setlanguage"}
TG_USER_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]{3,31}$")
TOKEN_SPLIT_RE = re.compile(r"[\s,;]+")

COLS = {
    "id": ["id организации", "id"],
    "name": ["название", "name"],
    "sites": ["сайты", "сайт", "website", "site"],
    "social": ["соцсети", "соцсеть", "social"],
    "yandex": ["ссылка на карточку", "ссылка яндекс", "yandex"],
    "category": ["основная категория", "категория"],
    "city": ["город"],
    "address": ["адрес"],
}


def host_of(token: str) -> str:
    t = token.strip()
    if not re.match(r"^[a-z]+://", t, re.I):
        t = "http://" + t
    try:
        h = (urlparse(t).hostname or "").lower()
    except ValueError:
        return ""
    return h[4:] if h.startswith("www.") else h


def is_real_site(cell: str) -> bool:
    for tok in TOKEN_SPLIT_RE.split(cell or ""):
        if not tok or "." not in tok:
            continue
        h = host_of(tok)
        if h and h not in NOT_A_SITE and not any(h.endswith("." + d) for d in NOT_A_SITE):
            return True
    return False


def extract_contacts(*cells):
    found, seen = [], set()
    for cell in cells:
        for tok in TOKEN_SPLIT_RE.split(cell or ""):
            if not tok:
                continue
            tok = tok.strip("()<>\"'")
            h = host_of(tok)
            if h in ("t.me", "telegram.me", "telegram.dog"):
                t = tok if re.match(r"^[a-z]+://", tok, re.I) else "https://" + tok
                path = [p for p in urlparse(t).path.split("/") if p]
                if not path:
                    continue
                user = path[0]
                if user.startswith("+") or user.lower() in TG_SKIP or not TG_USER_RE.match(user):
                    continue
                url = f"https://t.me/{user}"
                if url.lower() not in seen:
                    seen.add(url.lower())
                    found.append({"type": "tg", "url": url})
            elif h == "max.ru":
                t = tok if re.match(r"^[a-z]+://", tok, re.I) else "https://" + tok
                p = urlparse(t)
                path = [x for x in p.path.split("/") if x]
                if not path or path[0].lower() == "join" or path[0].lstrip("+").isdigit():
                    continue
                url = f"https://max.ru{p.path}".rstrip("/")
                if url.lower() not in seen:
                    seen.add(url.lower())
                    found.append({"type": "max", "url": url})
    found.sort(key=lambda c: 0 if c["type"] == "tg" else 1)
    return found


def read_rows(filename: str, data: bytes):
    name = filename.lower()
    if name.endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        ws = wb.active
        rows = [["" if v is None else str(v) for v in r] for r in ws.iter_rows(values_only=True)]
    elif name.endswith((".csv", ".txt")):
        text = None
        for enc in ("utf-8-sig", "cp1251"):
            try:
                text = data.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError("Не удалось прочитать кодировку файла")
        try:
            dialect = csv.Sniffer().sniff(text[:5000], delimiters=";,\t")
            delim = dialect.delimiter
        except csv.Error:
            delim = ";"
        rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=delim))
    else:
        raise ValueError("Нужен файл .csv или .xlsx")
    rows = [r for r in rows if any(c.strip() for c in r)]
    if not rows:
        raise ValueError("Файл пустой")
    return rows[0], rows[1:]


def find_col(header, keys):
    low = [h.strip().lower() for h in header]
    for k in keys:
        if k in low:
            return low.index(k)
    return None


def process(filename: str, data: bytes):
    header, body = read_rows(filename, data)
    idx = {k: find_col(header, v) for k, v in COLS.items()}
    if idx["name"] is None:
        raise ValueError("Не нашёл колонку «Название»")

    def get(row, key):
        i = idx[key]
        return row[i].strip() if i is not None and i < len(row) else ""

    stats = {"total": 0, "with_site": 0, "no_contact": 0, "duplicates": 0}
    items, seen_ids = [], set()
    for row in body:
        name = get(row, "name")
        if not name:
            continue
        stats["total"] += 1
        oid = get(row, "id")
        if oid and oid in seen_ids:
            stats["duplicates"] += 1
            continue
        seen_ids.add(oid)

        if is_real_site(get(row, "sites")):
            stats["with_site"] += 1
            continue
        contacts = extract_contacts(get(row, "social"), get(row, "sites"))
        if not contacts:
            stats["no_contact"] += 1
            continue
        yandex = get(row, "yandex") or (f"https://yandex.ru/maps/org/{oid}" if oid else "")
        items.append({
            "id": oid, "name": name, "contacts": contacts, "yandex": yandex,
            "category": get(row, "category"), "city": get(row, "city"),
        })
    stats["selected"] = len(items)
    return {"items": items, "stats": stats}


# ====================== СОСТОЯНИЕ СЕНДЕРА ======================
SEND_STATE = {
    "running": False, "paused": False, "stop_requested": False,
    "current_index": -1, "total": 0,
    "sent": 0, "failed": 0, "skipped": 0,
    "log": [], "sent_ids": set(),
}
_sender_thread = None
_state_lock = threading.Lock()


def _reset_state(total=0):
    with _state_lock:
        SEND_STATE.update({
            "running": True, "paused": False, "stop_requested": False,
            "current_index": -1, "total": total,
            "sent": 0, "failed": 0, "skipped": 0,
            "log": [], "sent_ids": set(),
        })


def _should_stop():
    with _state_lock:
        return SEND_STATE["stop_requested"]


def _wait_if_paused():
    while True:
        with _state_lock:
            paused = SEND_STATE["paused"]
        if not paused or _should_stop():
            break
        time.sleep(0.5)


def _log(text):
    with _state_lock:
        SEND_STATE["log"].append(text)
        if len(SEND_STATE["log"]) > 200:
            SEND_STATE["log"] = SEND_STATE["log"][-200:]


async def _send_one(client, item, message_text):
    tg_contact = next((c for c in item["contacts"] if c["type"] == "tg"), None)
    if not tg_contact:
        return False, "no_tg_contact"
    username = tg_contact["url"].rstrip("/").split("/")[-1]
    try:
        entity = await client.get_entity(username)
    except Exception as e:
        return False, f"entity_error: {type(e).__name__}"

    if not isinstance(entity, User):
        return False, "not_a_user"
    if getattr(entity, "bot", False):
        return False, "is_bot"
    if getattr(entity, "deleted", False):
        return False, "deleted_account"

    try:
        await client.send_message(entity, message_text)
        return True, "sent"
    except errors.FloodWaitError as e:
        return False, f"flood_wait_{e.seconds}s"
    except errors.UserPrivacyRestrictedError:
        return False, "privacy_restricted"
    except Exception as e:
        return False, f"send_error: {type(e).__name__}"


def _sender_worker(items, tpl, add_yandex):
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    async def _run():
        if not TELETHON_OK:
            _log("Telethon не установлен")
            return
        if not TG_API_ID or not TG_API_HASH:
            _log("Не заданы TG_API_ID / TG_API_HASH")
            return
        session_str = (os.environ.get("TG_SESSION_STRING") or _runtime_session_string or "").strip()
        if not session_str:
            _log("Нет сессии. Зайди в /setup и авторизуйся.")
            return

        client = TelegramClient(StringSession(session_str), int(TG_API_ID), TG_API_HASH)
        try:
            await client.connect()
            if not await client.is_user_authorized():
                _log("Сессия не авторизована. Пройди /setup заново.")
                return

            for idx, item in enumerate(items):
                if _should_stop():
                    _log("Остановлено пользователем")
                    break
                _wait_if_paused()
                if _should_stop():
                    break

                with _state_lock:
                    SEND_STATE["current_index"] = idx

                msg = tpl.replace("{name}", item["name"]).replace("{yandex}", item["yandex"] or "")
                if add_yandex and "{yandex}" not in tpl and item["yandex"]:
                    msg += "\n\nЯндекс.Карты: " + item["yandex"]

                ok, reason = await _send_one(client, item, msg)
                with _state_lock:
                    key = item["id"] or item["name"]
                    if ok:
                        SEND_STATE["sent"] += 1
                        SEND_STATE["sent_ids"].add(key)
                        _log(f"✅ {item['name']} — отправлено")
                    else:
                        if reason == "no_tg_contact":
                            SEND_STATE["skipped"] += 1
                        else:
                            SEND_STATE["failed"] += 1
                        _log(f"⏭ {item['name']} — {reason}")

                if idx < len(items) - 1:
                    waited = 0
                    while waited < SEND_DELAY and not _should_stop():
                        _wait_if_paused()
                        time.sleep(0.5)
                        waited += 0.5
        finally:
            await client.disconnect()

    try:
        loop.run_until_complete(_run())
    except Exception as e:
        _log(f"Критическая ошибка: {e}")
    finally:
        loop.close()
        with _state_lock:
            SEND_STATE["running"] = False
            SEND_STATE["paused"] = False


# ====================== API СЕНДЕРА ======================
@app.post("/send_start")
def send_start():
    global _sender_thread
    if not TELETHON_OK:
        return jsonify(error="Telethon не установлен на сервере"), 400
    data = request.get_json(silent=True) or {}
    items = data.get("items")
    tpl = data.get("tpl", "")
    if not items:
        return jsonify(error="Нет списка для отправки"), 400
    if not tpl.strip():
        return jsonify(error="Пустой шаблон сообщения"), 400
    with _state_lock:
        if SEND_STATE["running"]:
            return jsonify(error="Рассылка уже запущена"), 409
        _reset_state(total=len(items))
    _sender_thread = threading.Thread(target=_sender_worker,
                                      args=(items, tpl, bool(data.get("add_yandex"))),
                                      daemon=True)
    _sender_thread.start()
    return jsonify(ok=True, message="Рассылка запущена")


@app.post("/send_pause")
def send_pause():
    with _state_lock:
        if not SEND_STATE["running"]:
            return jsonify(error="Рассылка не запущена"), 400
        SEND_STATE["paused"] = True
    return jsonify(ok=True, paused=True)


@app.post("/send_resume")
def send_resume():
    with _state_lock:
        if not SEND_STATE["running"]:
            return jsonify(error="Рассылка не запущена"), 400
        SEND_STATE["paused"] = False
    return jsonify(ok=True, paused=False)


@app.post("/send_stop")
def send_stop():
    with _state_lock:
        SEND_STATE["stop_requested"] = True
        SEND_STATE["paused"] = False
    return jsonify(ok=True)


@app.get("/send_status")
def send_status():
    with _state_lock:
        return jsonify({
            "running": SEND_STATE["running"], "paused": SEND_STATE["paused"],
            "current_index": SEND_STATE["current_index"], "total": SEND_STATE["total"],
            "sent": SEND_STATE["sent"], "failed": SEND_STATE["failed"],
            "skipped": SEND_STATE["skipped"],
            "sent_ids": list(SEND_STATE["sent_ids"]),
            "log": SEND_STATE["log"][-30:],
        })


# ====================== SETUP (логин в TG через браузер) ======================
_login_clients = {}  # token -> {"loop", "client", "phone", "hash"}


def _check_setup(token):
    return bool(SETUP_TOKEN) and token == SETUP_TOKEN


@app.get("/setup")
def setup_page():
    token = request.args.get("token", "")
    if not _check_setup(token):
        return "Not found", 404
    return render_template_string(SETUP_PAGE, token=token)


@app.post("/setup/send_code")
def setup_send_code():
    data = request.get_json(silent=True) or {}
    if not _check_setup(data.get("token", "")):
        return jsonify(error="forbidden"), 403
    if not TELETHON_OK:
        return jsonify(error="Telethon не установлен"), 400
    if not TG_API_ID or not TG_API_HASH:
        return jsonify(error="TG_API_ID / TG_API_HASH не заданы на Render"), 400
    phone = (data.get("phone") or "").strip()
    if not phone:
        return jsonify(error="Укажи телефон, например +79001234567"), 400
    token = data["token"]

    # закрываем предыдущий клиент, если был
    old = _login_clients.pop(token, None)
    if old:
        try:
            old["loop"].run_until_complete(old["client"].disconnect())
        except Exception:
            pass
        try:
            old["loop"].close()
        except Exception:
            pass

    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        client = TelegramClient(StringSession(), int(TG_API_ID), TG_API_HASH, loop=loop)
        loop.run_until_complete(client.connect())
        sent = loop.run_until_complete(
            client.send_code_request(phone, force_sms=bool(data.get("force_sms")))
        )
        _login_clients[token] = {
            "loop": loop, "client": client,
            "phone": phone, "hash": sent.phone_code_hash,
        }
        return jsonify(ok=True, message="Код отправлен")
    except Exception as e:
        return jsonify(error=f"{type(e).__name__}: {e}"), 400


@app.post("/setup/verify")
def setup_verify():
    global _runtime_session_string
    data = request.get_json(silent=True) or {}
    token = data.get("token", "")
    if not _check_setup(token):
        return jsonify(error="forbidden"), 403
    state = _login_clients.get(token)
    if not state:
        return jsonify(error="Сначала запроси код"), 400

    code = (data.get("code") or "").strip()
    password = (data.get("password") or "").strip()
    loop, client = state["loop"], state["client"]

    try:
        if password:
            loop.run_until_complete(client.sign_in(password=password))
        else:
            if not code:
                return jsonify(error="Укажи код"), 400
            try:
                loop.run_until_complete(client.sign_in(
                    phone=state["phone"], code=code, phone_code_hash=state["hash"]
                ))
            except errors.SessionPasswordNeededError:
                return jsonify(needs_password=True,
                               message="Включена 2FA — введи облачный пароль")
    except errors.PhoneCodeInvalidError:
        return jsonify(error="Неверный код"), 400
    except errors.PhoneCodeExpiredError:
        return jsonify(error="Код истёк — запроси заново"), 400
    except errors.PasswordHashInvalidError:
        return jsonify(error="Неверный пароль 2FA"), 400
    except Exception as e:
        return jsonify(error=f"{type(e).__name__}: {e}"), 400

    if not loop.run_until_complete(client.is_user_authorized()):
        return jsonify(error="Не удалось авторизоваться"), 400

    session_str = client.session.save()
    _runtime_session_string = session_str
    try:
        loop.run_until_complete(client.disconnect())
    except Exception:
        pass
    try:
        loop.close()
    except Exception:
        pass
    _login_clients.pop(token, None)
    return jsonify(ok=True, session=session_str)


# ====================== ГЛАВНАЯ ======================
@app.post("/upload")
def upload():
    f = request.files.get("file")
    if not f:
        return jsonify(error="Файл не выбран"), 400
    try:
        return jsonify(process(f.filename or "", f.read()))
    except Exception as e:
        return jsonify(error=str(e)), 400


@app.get("/")
def index():
    return render_template_string(PAGE)


# ====================== HTML: SETUP ======================
SETUP_PAGE = r"""<!doctype html><html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Вход в Telegram</title>
<style>
body{font:15px/1.5 system-ui,Segoe UI,Roboto,sans-serif;padding:20px;max-width:620px;margin:0 auto;background:#f6f7f9;color:#1b1f24}
h1{font-size:20px}
label{display:block;margin-top:14px;font-weight:600;font-size:14px}
input,button{font:inherit;width:100%;box-sizing:border-box;padding:10px;border-radius:8px;border:1px solid #ccc;margin-top:4px}
button{cursor:pointer;background:#2f6fed;color:#fff;border:0}
button:hover{filter:brightness(.95)}
.box{background:#fff;padding:12px;border-radius:8px;margin:12px 0;word-break:break-all;font-family:ui-monospace,monospace;font-size:12px;border:1px solid #e2e5ea}
#log{margin-top:14px;font-size:14px}
#result{display:none;margin-top:20px}
.hint{color:#6b7480;font-size:13px}
</style></head><body>
<h1>Первичный вход в Telegram</h1>
<p class="hint">Введи телефон в формате <code>+79001234567</code>. Код придёт в Telegram (или по SMS, если отметить галочку). Если включена 2FA — попросит пароль. VPN не нужен: авторизация идёт с серверов Render.</p>

<label>Телефон</label>
<input id="phone" type="tel" placeholder="+79001234567">
<label style="margin-top:10px;font-weight:400;font-size:13px">
  <input type="checkbox" id="smsgate" style="width:auto;margin-right:6px"> Отправить код по SMS
</label>
<button onclick="sendCode()">Отправить код</button>

<div id="step2" style="display:none">
  <label>Код из Telegram/SMS</label>
  <input id="code" type="text" placeholder="12345">
  <button onclick="verify(false)">Проверить код</button>
</div>

<div id="step3" style="display:none">
  <label>Облачный пароль (2FA)</label>
  <input id="pwd" type="password">
  <button onclick="verify(true)">Войти</button>
</div>

<div id="log"></div>

<div id="result">
  <h2>Готово ✅</h2>
  <p>Скопируй строку ниже → вставь её в Render → <b>Environment</b> → <code>TG_SESSION_STRING</code> → Manual Deploy.</p>
  <div class="box" id="sess"></div>
</div>

<script>
const TOKEN = new URLSearchParams(location.search).get('token');
const $ = id => document.getElementById(id);
function log(msg, err){ $('log').innerHTML = `<span style="color:${err?'#c00':'#1f9d55'}">${msg}</span>`; }
async function api(path, body){
  const r = await fetch(path, {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({token: TOKEN, ...body})});
  return r.json();
}
async function sendCode(){
  log('Отправляю...');
  const j = await api('/setup/send_code', {phone: $('phone').value.trim(), force_sms: $('smsgate').checked});
  if (j.error){ log(j.error, true); return; }
  log('Код отправлен. Проверь Telegram или SMS.');
  $('step2').style.display = 'block';
}
async function verify(usePwd){
  log('Проверяю...');
  const body = usePwd ? {password: $('pwd').value} : {code: $('code').value.trim()};
  const j = await api('/setup/verify', body);
  if (j.needs_password){ $('step3').style.display = 'block'; log('Введи пароль 2FA'); return; }
  if (j.error){ log(j.error, true); return; }
  if (j.ok){
    log('Авторизован!');
    $('step2').style.display = 'none';
    $('step3').style.display = 'none';
    $('result').style.display = 'block';
    $('sess').textContent = j.session;
  }
}
</script></body></html>"""


# ====================== HTML: ГЛАВНАЯ ======================
PAGE = r"""<!doctype html>
<html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Холодные сообщения — бизнесы без сайта</title>
<style>
:root{--bg:#f6f7f9;--card:#fff;--tx:#1b1f24;--mut:#6b7480;--bd:#e2e5ea;--ac:#2f6fed;--ok:#1f9d55;--warn:#e6a817;--danger:#d33}
@media(prefers-color-scheme:dark){:root{--bg:#14171a;--card:#1d2125;--tx:#e8eaed;--mut:#9aa3ad;--bd:#2d3339;--ac:#6b9bff;--ok:#4cc27f;--warn:#f0b400;--danger:#ff6b6b}}
*{box-sizing:border-box}body{margin:0;font:15px/1.5 system-ui,Segoe UI,Roboto,sans-serif;background:var(--bg);color:var(--tx)}
.wrap{max-width:960px;margin:0 auto;padding:24px 16px 80px}
h1{font-size:22px;margin:0 0 16px}
.card{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:16px;margin-bottom:16px}
#drop{border:2px dashed var(--bd);border-radius:12px;padding:28px;text-align:center;cursor:pointer;color:var(--mut)}
#drop.over{border-color:var(--ac);color:var(--ac)}
textarea{width:100%;font:inherit;color:inherit;background:transparent;border:1px solid var(--bd);border-radius:8px;padding:10px;resize:vertical}
button,.btn{font:inherit;border:1px solid var(--bd);background:var(--card);color:var(--tx);border-radius:8px;padding:7px 12px;cursor:pointer;text-decoration:none;display:inline-block}
button.pri{background:var(--ac);border-color:var(--ac);color:#fff}
button.ok{background:var(--ok);border-color:var(--ok);color:#fff}
button.warn{background:var(--warn);border-color:var(--warn);color:#000}
button.danger{background:var(--danger);border-color:var(--danger);color:#fff}
button:disabled{opacity:.5;cursor:not-allowed}
.row{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
.mut{color:var(--mut);font-size:13px}
.stats span{margin-right:14px}
.item.done{opacity:.55}
.item h3{margin:0 0 4px;font-size:16px}
.tag{font-size:12px;padding:2px 8px;border-radius:99px;border:1px solid var(--bd);color:var(--mut)}
.msg{margin:10px 0;white-space:pre-wrap;background:var(--bg);border-radius:8px;padding:10px;font-size:14px}
#err{color:var(--danger);margin-top:8px}
#send-log{max-height:220px;overflow:auto;font-size:13px;background:var(--bg);border-radius:8px;padding:8px;margin-top:8px}
#send-log div{padding:1px 0}
.badge{display:inline-block;font-size:12px;padding:1px 7px;border-radius:99px;background:var(--bd);color:var(--mut)}
.badge.run{background:#e6f0ff;color:#2f6fed}
.badge.pause{background:#fff3cd;color:#7a5c00}
</style></head><body><div class="wrap">
<h1>Бизнесы без сайта → готовые сообщения</h1>

<div class="card">
  <div id="drop">Перетащи сюда файл <b>.csv</b> или <b>.xlsx</b> или нажми, чтобы выбрать</div>
  <input id="file" type="file" accept=".csv,.xlsx,.xlsm,.txt" hidden>
  <div id="err"></div>
</div>

<div class="card">
  <div class="row" style="justify-content:space-between"><b>Шаблон сообщения</b>
    <span class="mut">{name} — название, {yandex} — Яндекс.Карты</span></div>
  <textarea id="tpl" rows="9"></textarea>
  <div class="row" style="margin-top:8px">
    <label><input type="checkbox" id="addYa"> Добавлять ссылку на Яндекс.Карты, если {yandex} нет в шаблоне</label>
    <button id="resetTpl">Сбросить шаблон</button>
  </div>
</div>

<div id="result" hidden>
  <div class="card row" style="justify-content:space-between">
    <div class="stats" id="stats"></div>
    <div class="row">
      <label><input type="checkbox" id="hideDone"> Скрыть отмеченные</label>
      <button id="dl">Скачать .txt</button>
    </div>
  </div>

  <div class="card">
    <div class="row" style="justify-content:space-between;margin-bottom:10px">
      <b>Отправка в Telegram</b>
      <span id="send-badge" class="badge">не запущено</span>
    </div>
    <div class="row">
      <button id="btnStart" class="pri">Отправить в Telegram</button>
      <button id="btnPause" class="warn" disabled>Пауза</button>
      <button id="btnResume" class="ok" disabled>Продолжить</button>
      <button id="btnStop" class="danger" disabled>Остановить</button>
      <span class="mut" id="send-progress"></span>
    </div>
    <div id="send-log"></div>
  </div>

  <div id="list"></div>
</div>
</div>

<script>
const DEFAULT_TPL = 'Здравствуйте! Я Арсений, заметил, что у «{name}» пока нет сайта — а сейчас многие клиенты перед обращением гуглят или смотрят Яндекс.Карты. Без сайта часть из них просто уходит к конкурентам, у которых он есть. По данным исследований, наличие сайта увеличивает количество клиентов в среднем на 30–40%. Я делаю простые, но красивые сайты под ключ — с отзывами, ценами и фото. Стоит недорого, делаю быстро (в этом деле 4 года, знаю, что реально работает для такого бизнеса). Было бы интересно сделать для вас сайт? Могу предоставить вам свою последнюю работу';
const $ = id => document.getElementById(id);
let items = [];
const store = {
  get(k, d){ try{ const v = localStorage.getItem(k); return v===null? d : JSON.parse(v);}catch(e){return d;} },
  set(k, v){ try{ localStorage.setItem(k, JSON.stringify(v)); }catch(e){} }
};
let done = new Set(store.get('done', []));
$('tpl').value = store.get('tpl', DEFAULT_TPL);
$('addYa').checked = store.get('addYa', false);

function msgFor(it){
  let t = $('tpl').value.replaceAll('{name}', it.name).replaceAll('{yandex}', it.yandex || '');
  if ($('addYa').checked && !$('tpl').value.includes('{yandex}') && it.yandex)
    t += '\n\nЯндекс.Карты: ' + it.yandex;
  return t;
}
const label = c => c.type === 'tg' ? 'Telegram' : 'MAX';

function render(){
  const list = $('list'); list.innerHTML = '';
  const hide = $('hideDone').checked;
  items.forEach((it, i) => {
    const key = it.id || it.name;
    const isDone = done.has(key);
    if (hide && isDone) return;
    const d = document.createElement('div');
    d.className = 'card item' + (isDone ? ' done' : '');

    const h = document.createElement('h3'); h.textContent = `${i+1}. ${it.name}`;
    const meta = document.createElement('div'); meta.className = 'row mut';
    [it.city, it.category].filter(Boolean).forEach(x => {
      const s = document.createElement('span'); s.className = 'tag'; s.textContent = x; meta.append(s);
    });

    const links = document.createElement('div'); links.className = 'row'; links.style.marginTop = '8px';
    it.contacts.forEach(c => {
      const a = document.createElement('a'); a.className = 'btn'; a.href = c.url; a.target = '_blank'; a.rel = 'noopener';
      a.textContent = `${label(c)}: ${c.url.replace('https://','')}`; links.append(a);
    });
    if (it.yandex){
      const a = document.createElement('a'); a.className = 'btn'; a.href = it.yandex; a.target = '_blank'; a.rel = 'noopener';
      a.textContent = 'Яндекс.Карты'; links.append(a);
    }

    const m = document.createElement('div'); m.className = 'msg'; m.textContent = msgFor(it);

    const act = document.createElement('div'); act.className = 'row';
    const cp = document.createElement('button'); cp.className = 'pri'; cp.textContent = 'Копировать текст';
    cp.onclick = async () => {
      try { await navigator.clipboard.writeText(msgFor(it)); }
      catch(e){ const ta = document.createElement('textarea'); ta.value = msgFor(it); document.body.append(ta); ta.select(); document.execCommand('copy'); ta.remove(); }
      cp.textContent = 'Скопировано ✓'; setTimeout(() => cp.textContent = 'Копировать текст', 1500);
    };
    const lb = document.createElement('label');
    const cb = document.createElement('input'); cb.type = 'checkbox'; cb.checked = isDone;
    cb.onchange = () => { cb.checked ? done.add(key) : done.delete(key); store.set('done', [...done]); render(); };
    lb.append(cb, ' Написал(а)');
    act.append(cp, lb);

    d.append(h, meta, links, m, act); list.append(d);
  });
}

function upload(file){
  $('err').textContent = '';
  const fd = new FormData(); fd.append('file', file);
  fetch('/upload', {method:'POST', body: fd}).then(r => r.json()).then(j => {
    if (j.error){ $('err').textContent = j.error; return; }
    items = j.items; const s = j.stats;
    $('stats').innerHTML = `<span>В файле: <b>${s.total}</b></span><span>С сайтом: <b>${s.with_site}</b></span>` +
      `<span>Без сайта и без TG/MAX: <b>${s.no_contact}</b></span><span style="color:var(--ok)">Подходит: <b>${s.selected}</b></span>`;
    $('result').hidden = false; render();
  }).catch(e => $('err').textContent = 'Ошибка: ' + e);
}

function setSendUI(st){
  const running = st.running, paused = st.paused;
  $('btnStart').disabled = running;
  $('btnPause').disabled = !running || paused;
  $('btnResume').disabled = !running || !paused;
  $('btnStop').disabled = !running;
  const badge = $('send-badge');
  badge.className = 'badge' + (running ? (paused ? ' pause' : ' run') : '');
  badge.textContent = running ? (paused ? 'пауза' : 'работает') : 'не запущено';
  $('send-progress').textContent = running
    ? `Обработано: ${st.current_index+1} / ${st.total} · отправлено ${st.sent} · ошибок ${st.failed} · пропущено ${st.skipped}`
    : '';
  $('send-log').innerHTML = (st.log || []).map(x => `<div>${x}</div>`).join('');
  if (st.sent_ids && st.sent_ids.length){
    st.sent_ids.forEach(id => done.add(id));
    store.set('done', [...done]);
    render();
  }
}
function pollSend(){ fetch('/send_status').then(r => r.json()).then(setSendUI).catch(()=>{}); }
setInterval(pollSend, 2000);
pollSend();

$('btnStart').onclick = () => {
  if (!items.length){ alert('Сначала загрузи файл'); return; }
  fetch('/send_start', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({items, tpl: $('tpl').value, add_yandex: $('addYa').checked})
  }).then(r => r.json()).then(j => { if (j.error) alert(j.error); pollSend(); });
};
$('btnPause').onclick = () => fetch('/send_pause', {method:'POST'}).then(pollSend);
$('btnResume').onclick = () => fetch('/send_resume', {method:'POST'}).then(pollSend);
$('btnStop').onclick = () => { if(confirm('Остановить рассылку?')) fetch('/send_stop', {method:'POST'}).then(pollSend); };

$('drop').onclick = () => $('file').click();
$('file').onchange = e => e.target.files[0] && upload(e.target.files[0]);
['dragover','dragenter'].forEach(ev => $('drop').addEventListener(ev, e => { e.preventDefault(); $('drop').classList.add('over'); }));
['dragleave','drop'].forEach(ev => $('drop').addEventListener(ev, e => { e.preventDefault(); $('drop').classList.remove('over'); }));
$('drop').addEventListener('drop', e => e.dataTransfer.files[0] && upload(e.dataTransfer.files[0]));
$('tpl').oninput = () => { store.set('tpl', $('tpl').value); render(); };
$('addYa').onchange = () => { store.set('addYa', $('addYa').checked); render(); };
$('hideDone').onchange = render;
$('resetTpl').onclick = () => { $('tpl').value = DEFAULT_TPL; store.set('tpl', DEFAULT_TPL); render(); };
$('dl').onclick = () => {
  const txt = items.map((it, i) => `${i+1}. ${it.name} (ID ${it.id}) — ${it.contacts.map(c => c.url).join(' | ')}` +
    (it.yandex ? `\nЯндекс.Карты: ${it.yandex}` : '') + `\n${msgFor(it)}\n`).join('\n');
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([txt], {type:'text/plain;charset=utf-8'}));
  a.download = 'messages.txt'; a.click();
};
</script></body></html>"""


def free_port(start=5000):
    for p in range(start, start + 50):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return start


if __name__ == "__main__":
    port = free_port()
    url = f"http://127.0.0.1:{port}"
    print("\n" + "=" * 50)
    print(f"  Готово! Открой в браузере: {url}")
    print("  Остановить программу: Ctrl+C")
    print("=" * 50 + "\n")
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=port, debug=False)
