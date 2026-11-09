"""
SupplierHub — база поставщиков и прайс-листов.
Работает как локальный веб-сервер: открывается и на компьютере, и на iPhone
(через Wi-Fi) — без установки приложения из App Store.

Запуск:  python server.py
Открыть: http://localhost:8000           (на компьютере)
         http://<IP-адрес>:8000          (на iPhone в той же Wi-Fi сети)

Зависимости: только стандартная библиотека Python + openpyxl (для Excel).
"""

import os
import re
import sys
import json
import html
import shutil
import socket
import sqlite3
import subprocess
import urllib.parse
from datetime import datetime
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "supplierhub.db")
UP_DIR = os.path.join(APP_DIR, "uploads")
PORT = int(os.environ.get("PORT") or os.environ.get("SH_PORT") or "8000")
RATE = 16300.0  # сум за 1 USD
os.makedirs(UP_DIR, exist_ok=True)

PRICE_RE = re.compile(r"(\d[\d\s.,]{0,14})\s*(usd|uzs|сум|\$|cny|rmb|юан|元|руб)", re.I)
HEADER_HINTS = ("наимен", "товар", "описан", "product", "item", "name", "商品", "名称", "описан")
PRICE_HINTS = ("цена", "стоим", "прайс", "price", "cost", "价格", "单价", "сумм")


# ------------------------- база -------------------------
def db():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.execute("""CREATE TABLE IF NOT EXISTS suppliers(
        id INTEGER PRIMARY KEY AUTOINCREMENT, created TEXT, name TEXT, country TEXT,
        contact TEXT, phone TEXT, email TEXT, wechat TEXT, note TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS uploads(
        id INTEGER PRIMARY KEY AUTOINCREMENT, created TEXT, supplier_id INTEGER,
        filename TEXT, items INTEGER, hash TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS settings(
        key TEXT PRIMARY KEY, value TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS kp_history(
        id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER, created TEXT,
        markup REAL, total_buy REAL, total_sell REAL, path TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS orders(
        id INTEGER PRIMARY KEY AUTOINCREMENT, created TEXT, filename TEXT,
        markup REAL, total_buy REAL, total_sell REAL)""")
    con.execute("""CREATE TABLE IF NOT EXISTS order_items(
        id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER, req_name TEXT,
        qty REAL, match_name TEXT, supplier_id INTEGER, price REAL, currency TEXT, usd REAL)""")
    con.execute("""CREATE TABLE IF NOT EXISTS items(
        id INTEGER PRIMARY KEY AUTOINCREMENT, upload_id INTEGER, supplier_id INTEGER,
        name TEXT, name_lc TEXT, price REAL, currency TEXT, unit TEXT, raw TEXT)""")
    # миграция: колонка name_lc (SQLite lower() не работает с кириллицей)
    ucols = [r[1] for r in con.execute("PRAGMA table_info(uploads)")]
    if "hash" not in ucols:
        con.execute("ALTER TABLE uploads ADD COLUMN hash TEXT")
    cols = [r[1] for r in con.execute("PRAGMA table_info(items)")]
    if "name_lc" not in cols:
        con.execute("ALTER TABLE items ADD COLUMN name_lc TEXT")
        con.execute("UPDATE items SET name_lc = lower(name)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_items_lc ON items(name_lc)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_items_sup ON items(supplier_id)")
    con.commit()
    return con



def to_number(text):
    """Аккуратно превращает строку в число. Даты и мусор возвращают None."""
    if text is None:
        return None
    if isinstance(text, (int, float)):
        return float(text)
    t = str(text)
    # отбрасываем похожее на дату: 26.10.2026 / 2026-10-26
    if re.fullmatch(r"\s*\d{1,4}[.\-/]\d{1,2}[.\-/]\d{2,4}\s*", t):
        return None
    t = re.sub(r"[^\d.,]", "", t).strip(".,")
    if not t:
        return None
    # если несколько разделителей — последний считаем десятичным
    if t.count(".") > 1 or t.count(",") > 1:
        t = t.replace(",", "")
        head, _, tail = t.rpartition(".")
        t = head.replace(".", "") + ("." + tail if tail else "")
    else:
        t = t.replace(",", ".")
    try:
        v = float(t)
    except ValueError:
        return None
    return v if v > 0 else None



def parse_screenshot(path):
    """Распознаёт прайс со скриншота через Windows OCR (shot_to_price.py)."""
    try:
        import importlib.util as iu
        script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                              "Documents", "deepseek-harness", "default-workspace", "shot_to_price.py")
        script = os.path.abspath(script)
        if not os.path.exists(script):
            return []
        spec = iu.spec_from_file_location("shot_to_price", script)
        mod = iu.module_from_spec(spec)
        spec.loader.exec_module(mod)
        items = mod.parse_image(path, "скриншот", write=False)
        return [(it["name"], it["price"], it["currency"], "шт") for it in items]
    except Exception as e:
        print("OCR error:", e)
        return []



def read_pdf_text(path):
    """Текст из PDF: сначала pypdf (работает везде), иначе pdftotext (Windows)."""
    text = ""
    try:
        import pypdf
        reader = pypdf.PdfReader(path)
        parts = []
        for page in reader.pages[:40]:
            try:
                parts.append(page.extract_text() or "")
            except Exception:
                continue
        text = "\n".join(parts)
    except Exception:
        text = ""
    if text.strip():
        return text
    pdftotext = shutil.which("pdftotext")
    if pdftotext:
        tmp = path + ".txt"
        try:
            subprocess.run([pdftotext, "-layout", "-enc", "UTF-8", path, tmp],
                           capture_output=True, timeout=180)
            if os.path.exists(tmp):
                with open(tmp, encoding="utf-8", errors="ignore") as f:
                    text = f.read()
                os.remove(tmp)
        except Exception:
            pass
    return text


def preprocess_image(path):
    """Увеличиваем и усиливаем контраст — OCR читает заметно лучше."""
    try:
        from PIL import Image, ImageOps
        im = Image.open(path).convert("L")
        im = ImageOps.autocontrast(im, cutoff=2)
        w, h = im.size
        scale = 2 if max(w, h) < 1600 else 1
        if scale > 1:
            im = im.resize((w * scale, h * scale), Image.LANCZOS)
        out = path + "_prep.png"
        im.save(out)
        return out
    except Exception:
        return path


def ocr_image_text(path):
    """Текст со скриншота: Windows OCR (локально), иначе Tesseract (на хостинге)."""
    prep = preprocess_image(path)
    try:
        return _ocr_raw(prep)
    finally:
        if prep != path and os.path.exists(prep):
            try:
                os.remove(prep)
            except OSError:
                pass


def _ocr_raw(path):
    # 1) встроенный OCR Windows
    if os.name == "nt":
        try:
            ps = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                              "Documents", "deepseek-harness", "default-workspace", "ocr_json.ps1")
            ps = os.path.abspath(ps)
            if os.path.exists(ps):
                out = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                                      "-File", ps, "-Path", path],
                                     capture_output=True, text=True, encoding="utf-8", timeout=300)
                txt = (out.stdout or "").strip()
                if txt.startswith("{"):
                    import json as _json
                    data = _json.loads(txt)
                    return "\n".join(l.get("text", "") for l in data.get("lines", []))
        except Exception as e:
            print("windows ocr failed:", e)
    # 2) Tesseract (Linux/хостинг или установлен в Windows)
    tess = shutil.which("tesseract")
    if tess:
        try:
            out = subprocess.run([tess, path, "stdout", "-l", "rus+eng", "--psm", "6"],
                                 capture_output=True, text=True, encoding="utf-8", timeout=300)
            return out.stdout or ""
        except Exception as e:
            print("tesseract failed:", e)
    return ""


def items_from_text(text):
    """Из произвольного текста достаём строки «наименование ... цена»."""
    out = []
    for line in (text or "").splitlines():
        line = line.strip()
        if len(line) < 5:
            continue
        m = PRICE_RE.search(line)
        if not m:
            continue
        price = to_number(m.group(1))
        if not price:
            continue
        name = line[:m.start()].strip(" .,-—|;|\t")
        if len(name) < 3:
            continue
        out.append((name[:300], price, m.group(2).lower(), ""))
    return out


def parse_price_file(path):
    """Возвращает список (name, price, currency, unit). Поддержка xlsx, pdf, csv/txt."""
    ext = os.path.splitext(path)[1].lower()
    out = []

    if ext in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True, read_only=True)
        sheets = wb.worksheets
        # если в книге есть сводный лист — читаем только его (иначе посчитаем позиции дважды)
        for pref in ("Единый прайс", "Прайс", "Price"):
            for cand in sheets:
                if cand.title.strip().lower() == pref.lower():
                    sheets = [cand]
                    break
            if len(sheets) == 1:
                break
        for ws in sheets:
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            # ищем строку заголовков
            hdr_idx, name_col, price_col = 0, None, None
            for i, row in enumerate(rows[:25]):
                raw_cells = [str(c or "").strip() for c in row]
                nonempty = [c for c in raw_cells if c]
                # строка заголовков: минимум 3 непустых ячейки и короткие подписи
                # (иначе подзаголовок «...Курс 1 USD = 16 300 сум. Цена хранится...» ловится как шапка)
                if len(nonempty) < 3 or any(len(c) > 60 for c in nonempty):
                    continue
                cells = [c.lower() for c in raw_cells]
                ni = next((j for j, c in enumerate(cells) if any(h in c for h in HEADER_HINTS)), None)
                pi = next((j for j, c in enumerate(cells) if any(h in c for h in PRICE_HINTS)), None)
                if pi is not None:
                    hdr_idx, name_col, price_col = i, ni, pi
                    break
            for row in rows[hdr_idx + 1:]:
                if not row:
                    continue
                vals = list(row)
                name = ""
                if name_col is not None and name_col < len(vals) and vals[name_col]:
                    name = str(vals[name_col]).strip()
                else:
                    texts = [str(v).strip() for v in vals
                             if isinstance(v, str) and str(v).strip()
                             and ".xlsx" not in str(v).lower()
                             and ".pdf" not in str(v).lower()
                             and " / " not in str(v)]
                    name = max(texts, key=len) if texts else ""
                if not name or len(name) < 3:
                    continue
                price = None
                if price_col is not None and price_col < len(vals):
                    v = vals[price_col]
                    if isinstance(v, (int, float)):
                        price = float(v)
                    elif v is not None:
                        m = re.search(r"\d[\d\s.,]*", str(v))
                        if m:
                            price = to_number(m.group(0))
                if price is None:
                    line = " ".join(str(v) for v in vals if v is not None)
                    m = PRICE_RE.search(line)
                    if m:
                        price = to_number(m.group(1))
                if price is None:
                    continue
                raw_text = " ".join(str(v) for v in vals if v is not None)
                out.append((name[:300], price, guess_currency(price, raw_text), ""))
        wb.close()
        return out

    if ext == ".pdf":
        text = read_pdf_text(path)
        return items_from_text(text)

    # csv / txt
    with open(path, encoding="utf-8", errors="ignore") as f:
        for line in f:
            m = PRICE_RE.search(line)
            if not m:
                continue
            price = to_number(m.group(1))
            name = line[:m.start()].strip(" .,-—|;,|\t") or line[:120]
            if price and len(name.strip()) >= 3:
                out.append((name.strip()[:300], price, m.group(2).lower(), ""))
    return out


def guess_currency(price, text=""):
    cur = detect_currency(text) if text else "USD"
    if cur == "USD" and isinstance(price, (int, float)) and price >= 50000:
        return "UZS"
    return cur


def detect_currency(text):
    t = text.lower()
    if "sum" in t or "сум" in t or "uzs" in t:
        return "UZS"
    if "cny" in t or "rmb" in t or "юан" in t or "元" in t or "¥" in t:
        return "CNY"
    if "$" in t or "usd" in t:
        return "USD"
    return "USD"


# ------------------------- HTML -------------------------
CSS = """
:root { --bg:#10151c; --card:#18202b; --line:#2a3646; --txt:#e8eef6; --dim:#93a3b8; --acc:#3d8bfd; --ok:#2fbf71; }
* { box-sizing:border-box; -webkit-text-size-adjust:100%; }
body { margin:0; background:var(--bg); color:var(--txt);
       font:16px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif; }
header { background:#0b1016; border-bottom:1px solid var(--line); padding:14px 16px;
         position:sticky; top:0; z-index:5; }
h1 { font-size:18px; margin:0; }
nav a { color:var(--dim); text-decoration:none; margin-right:16px; font-size:15px; }
nav a.on { color:var(--acc); font-weight:600; }
main { padding:14px 16px 90px; max-width:1100px; margin:0 auto; }
.card { background:var(--card); border:1px solid var(--line); border-radius:12px; padding:14px; margin-bottom:14px; }
input, select, textarea, button { font:inherit; }
input[type=text], input[type=search], input[type=file], select, textarea {
    width:100%; padding:12px; border-radius:10px; border:1px solid var(--line);
    background:#0e141b; color:var(--txt); margin:6px 0; }
button, .btn { background:var(--acc); color:#fff; border:0; padding:12px 18px; border-radius:10px;
        font-weight:600; cursor:pointer; display:inline-block; text-decoration:none; }
.btn.gray { background:#243244; }
table { width:100%; border-collapse:collapse; font-size:14px; }
th, td { text-align:left; padding:8px 6px; border-bottom:1px solid var(--line); vertical-align:top; }
th { color:var(--dim); font-weight:600; font-size:13px; }
.p { color:var(--ok); font-weight:700; white-space:nowrap; }
.muted { color:var(--dim); font-size:13px; }
.badge { display:inline-block; background:#22303f; color:var(--dim); border-radius:6px; padding:2px 7px; font-size:12px; }
.row { display:flex; gap:10px; flex-wrap:wrap; }
.row > * { flex:1 1 180px; }
@media (max-width:640px){ h1{font-size:16px} main{padding:10px 10px 90px} }
"""


def page(title, body, active=""):
    nav = "".join(
        f'<a href="{href}" class="{"on" if active == href else ""}">{label}</a>'
        for href, label in (("/", "Поиск"), ("/order", "Заявка"), ("/kp", "КП"), ("/suppliers", "Поставщики"),
         ("/upload", "Прайс"), ("/changes", "Изменения"), ("/settings", "Реквизиты"))
    )
    return f"""<!DOCTYPE html><html lang="ru"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="SupplierHub">
<link rel="manifest" href="/manifest.json">
<link rel="apple-touch-icon" href="/icon.png">
<title>{html.escape(title)}</title><style>{CSS}</style></head>
<body><header><h1>SupplierHub</h1><nav>{nav}</nav></header><main>{body}</main></body></html>"""


def esc(v):
    return html.escape(str(v if v is not None else ""))


# ------------------------- страницы -------------------------

# ------------------------- умный поиск -------------------------
RU2LAT = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e", "ж": "zh",
    "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o",
    "п": "p", "р": "r", "с": "s", "т": "t", "у": "u", "ф": "f", "х": "h", "ц": "c",
    "ч": "ch", "ш": "sh", "щ": "sh", "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu",
    "я": "ya", "ў": "o", "қ": "q", "ғ": "g", "ҳ": "h", "ң": "n",
}
LAT2RU = {}
for _k, _v in RU2LAT.items():
    if _v and _v not in LAT2RU.values():
        LAT2RU.setdefault(_v, _k)


def _translit(text, table):
    return "".join(table.get(ch, ch) for ch in text)


def normalize(text):
    t = str(text or "").lower().replace("ё", "е").replace("’", "").replace("'", "")
    t = re.sub(r"[^0-9a-zа-яўқғҳң]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


SYNONYMS = {
    "коммутатор": ["switch"], "свитч": ["switch"], "маршрутизатор": ["router"],
    "роутер": ["router"], "принтер": ["printer"], "мфу": ["mfu", "printer", "printer"],
    "монитор": ["monitor"], "клавиатура": ["keyboard"], "мышь": ["mouse"],
    "наушники": ["headphone", "earphone"], "колонка": ["speaker"], "кабель": ["cable"],
    "картридж": ["cartridge", "toner"], "ноутбук": ["notebook", "laptop"],
    "телефон": ["phone", "smartphone"], "смартфон": ["smartphone"],
    "память": ["memory", "ram", "ddr"], "озу": ["ram", "ddr"], "ssd": ["ssd"],
    "жёсткий": ["hdd", "disk"], "блок": ["power", "supply"], "ибп": ["ups"],
    "кулер": ["cooler", "fan"], "корпус": ["case"], "сумка": ["bag"],
    "флешка": ["flash", "usb"], "камера": ["camera"], "мфу": ["mfp", "printer"],
    "сканер": ["scanner"], "проектор": ["projector"], "планшет": ["tablet"],
    "часы": ["watch"], "умные": ["smart"],
}

TRANSLIT_CACHE = {}


def query_variants(q):
    """Варианты запроса: как есть, латиницей, кириллицей, синонимы, без пробелов."""
    key = q
    if key in TRANSLIT_CACHE:
        return TRANSLIT_CACHE[key]
    base = normalize(q)
    lat = normalize(_translit(base, RU2LAT))
    cyr = normalize(_translit(base, LAT2RU)) if re.search(r"[a-z]", base) else base
    out = []
    # синонимы: каждое русское слово меняем на английский эквивалент
    syn_variants = []
    words = base.split()
    if words:
        for src_word, targets in SYNONYMS.items():
            if src_word in base:
                for t in targets:
                    syn_variants.append(base.replace(src_word, t))
    for v in (base, lat, cyr, base.replace(" ", ""), lat.replace(" ", "")) + tuple(syn_variants):
        if v and v not in out:
            out.append(v)
    TRANSLIT_CACHE[key] = out
    return out


def _lev(a, b):
    if a == b:
        return 0
    if not a or not b:
        return max(len(a), len(b))
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _token_similar(a, b):
    if a == b:
        return 1.0
    # короткие токены сравниваем только на равенство: иначе «b» совпадает со всем
    if len(a) < 3 or len(b) < 3:
        return 1.0 if a == b else 0.0
    if a in b or b in a:
        return 0.9
    d = _lev(a, b)
    return max(0.0, 1.0 - d / max(len(a), len(b)))


def score_name(name_norm, variants):
    """0..100 — насколько имя товара похоже на запрос."""
    if not name_norm:
        return 0.0
    best = 0.0
    name_tokens = name_norm.split()
    for qv in variants:
        q_tokens = qv.split()
        if not q_tokens:
            continue
        # полное вхождение строки запроса
        if qv in name_norm:
            best = max(best, 100.0 if len(qv) >= 4 else 80.0)
            continue
        # покрытие токенов запроса
        matched = 0.0
        for qt in q_tokens:
            best_t = max((_token_similar(qt, nt) for nt in name_tokens), default=0.0)
            matched += best_t
        cov = matched / len(q_tokens)
        # бонус за частичное вхождение подстроки
        if len(qv) >= 3 and any(len(nt) >= 3 and (qv in nt or nt in qv) for nt in name_tokens):
            cov = min(1.0, cov + 0.15)
        best = max(best, cov * 95.0)
    return best


_SEARCH_CACHE = {"ts": 0, "sig": None, "rows": None}


def _load_rows_cached():
    """Держим позиции в памяти: поиск в 5–10 раз быстрее обращений к SQLite."""
    import time as _time
    con = db()
    sig = con.execute("SELECT COUNT(*), COALESCE(MAX(id),0) FROM items").fetchone()
    sig = (sig[0], sig[1])
    con.close()
    now = _time.time()
    if _SEARCH_CACHE["rows"] is not None and _SEARCH_CACHE["sig"] == sig and now - _SEARCH_CACHE["ts"] < 60:
        return _SEARCH_CACHE["rows"]
    con = db()
    rows = con.execute("""SELECT i.id, i.name, i.price, i.currency, i.supplier_id, i.name_lc,
                                 s.name AS sup, s.country, s.contact, s.phone, s.email, s.wechat,
                                 u.filename
                          FROM items i JOIN suppliers s ON s.id=i.supplier_id
                          LEFT JOIN uploads u ON u.id=i.upload_id""").fetchall()
    con.close()
    norm_rows = [dict(r) for r in rows]
    for nr in norm_rows:
        nr["_nm"] = normalize(nr["name"])
        nr["_tokens"] = nr["_nm"].split()
    _SEARCH_CACHE.update({"ts": now, "sig": sig, "rows": rows, "norm": norm_rows})
    return rows


def invalidate_cache():
    _SEARCH_CACHE["rows"] = None


def smart_search(q, min_results=4, limit=120):
    _load_rows_cached()
    rows = _SEARCH_CACHE.get("norm") or []
    variants = query_variants(q)
    q_tokens = normalize(q).split()
    scored = []
    # быстрый предфильтр: точное вхождение хотя бы одного варианта запроса
    fast_variants = [v for v in variants if len(v) >= 3]
    q_tokens_all = [t for t in normalize(q).split() if len(t) >= 3]
    candidates = []
    for r in rows:
        nm = r.get("_nm") or normalize(r["name"])
        if not nm:
            continue
        if any(v in nm for v in fast_variants) or any(t[:4] in nm for t in q_tokens_all):
            candidates.append((r, nm))
    for r, nm in candidates:
        sc = score_name(nm, variants)
        name_tokens = r.get("_tokens") or nm.split()
        if q_tokens:
            hit = sum(1 for qt in q_tokens
                      if any(_token_similar(qt, nt) >= 0.6 for nt in name_tokens))
            cov = hit / len(q_tokens)
        else:
            cov = 1.0
        if sc <= 0:
            continue
        scored.append({
            "name": r["name"], "price": r["price"], "cur": r["currency"],
            "supplier": r["sup"], "country": r["country"],
            "contact": " · ".join(x for x in [r["contact"], r["phone"], r["email"], r["wechat"]] if x),
            "phone": r["phone"] or "", "email": r["email"] or "", "wechat": r["wechat"] or "",
            "person": r["contact"] or "",
            "file": r["filename"], "score": sc,
            "usd": (r["price"] or 0) if r["currency"] == "USD" else round((r["price"] or 0) / RATE, 2),
            "cov": cov,
            "lc": r["name_lc"], "sid": r["supplier_id"],
        })
    # строгий отбор: мусор со слабым совпадением не показываем вообще
    strong = [x for x in scored if x["score"] >= 75 and x.get("cov", 1) >= 0.6]
    if strong:
        strong.sort(key=lambda x: (-int(x["score"] // 10), x["usd"]))
        return strong[:limit], "точные совпадения · сначала самые дешёвые"
    # точных нет — показываем только близкие (не мусор)
    relaxed = [x for x in scored if x["score"] >= 55 and x.get("cov", 1) >= 0.65]
    if relaxed:
        relaxed.sort(key=lambda x: (-x["score"], x["usd"]))
        return relaxed[:limit], "приблизительные совпадения — уточните название"
    return [], "ничего не найдено"


def view_search(q=""):
    body = f"""<div class="card">
<form method="get" action="/"><input type="search" name="q" value="{esc(q)}"
 placeholder="Товар: коммутатор, router, H3C, кабель… (латиница или кириллица)"><button>Найти</button></form>
<div class="muted">Поиск нечёткий: учитывает раскладку, опечатки и части слов.</div></div>"""

    con = db()
    sups = con.execute("SELECT COUNT(*) c FROM suppliers").fetchone()["c"]
    items = con.execute("SELECT COUNT(*) c FROM items").fetchone()["c"]
    con.close()

    if q:
        rows, mode = smart_search(q)
        body += f"""<div class="card"><div class="muted">Режим: {esc(mode)} · найдено: {len(rows)}
 (минимум 4 результата по запросу)</div><table>
<tr><th>#</th><th>Товар</th><th>Совпадение</th><th>Цена</th><th>Поставщик</th><th>Контакт</th></tr>"""
        if not rows:
            body += '<tr><td colspan="5" class="muted">Ничего не найдено. Проверьте, загружен ли прайс этого поставщика.</td></tr>'
        # лучшая цена считается отдельно по каждой группе совпадения (нельзя
        # сравнивать картридж за 13$ с МФУ за 400$)
        tier_min = {}
        for r in rows:
            tier = int(r["score"] // 10)
            tier_min[tier] = min(tier_min.get(tier, r["usd"] or 1e12), r["usd"] or 1e12)
        overall_min = min((r["usd"] for r in rows), default=0)
        for idx, r in enumerate(rows, start=1):
            chg = last_change_for(r["sid"], r["lc"])
            delta = ""
            if chg:
                col = "#ff7b72" if chg["pct"] > 0 else "#2fbf71"
                sign = "+" if chg["pct"] > 0 else ""
                delta = f'<div style="color:{col};font-size:12px">{sign}{chg["pct"]:.1f}% к прошлому прайсу</div>'
            sc = r["score"]
            bar_col = "#2fbf71" if sc >= 80 else ("#d8a027" if sc >= 55 else "#8a94a6")
            tier = int(sc // 10)
            best_in_group = tier_min.get(tier, overall_min)
            is_cheapest = (r["usd"] == overall_min)
            row_style = ' style="background:#12331f"' if is_cheapest else ""
            badge = (' <span class="badge" style="background:#2fbf71;color:#08130c;font-weight:700">'
                     'самый дешёвый</span>') if is_cheapest else ""
            diff = ""
            if not is_cheapest and best_in_group:
                over = (r["usd"] - best_in_group) / best_in_group * 100
                if over > 1:
                    diff = f'<div class="muted">+{over:.0f}% к лучшей в группе</div>'
            body += (f'<tr{row_style}><td><b>{idx}</b></td>'
                     f'<td>{esc(r["name"])}<div class="muted">{esc(r["file"] or "")}</div></td>'
                     f'<td style="color:{bar_col};font-weight:700;white-space:nowrap">{sc:.0f}%</td>'
                     f'<td class="p">{esc(r["price"])} {esc(r["cur"])}{badge}{diff}{delta}</td>'
                     f'<td><a href="/supplier?id={r["sid"]}" style="color:#e8eef6">{esc(r["supplier"])}</a>'
                     f'<div class="muted">{esc(r["country"] or "")}</div></td>'
                     f'<td class="muted">{contact_links(r) or add_contact_link(r["sid"])}</td></tr>')
        body += "</table></div>"
    else:
        body += f'<div class="card"><div class="muted">Поставщиков: {sups} · позиций в прайсах: {items}. Введите запрос.</div></div>'
    return page("Поиск", body, "/")


def view_suppliers():
    con = db()
    rows = con.execute("SELECT * FROM suppliers ORDER BY id DESC").fetchall()
    con.close()
    body = """<div class="card"><form method="post" action="/suppliers/add">
<input type="text" name="name" placeholder="Компания *" required>
<div class="row">
<input type="text" name="country" placeholder="Страна">
<input type="text" name="contact" placeholder="Контактное лицо">
</div>
<div class="row">
<input type="text" name="phone" placeholder="Телефон / WhatsApp">
<input type="text" name="email" placeholder="E-mail">
<input type="text" name="wechat" placeholder="WeChat">
</div>
<input type="text" name="note" placeholder="Примечание">
<button>Добавить поставщика</button></form></div>"""
    body += '<div class="card"><table><tr><th>Компания</th><th>Страна</th><th>Контакт</th></tr>'
    for r in rows:
        c = " · ".join(x for x in [r["contact"], r["phone"], r["email"], r["wechat"]] if x)
        body += f"<tr><td>{esc(r['name'])}<div class='muted'>{esc(r['note'] or '')}</div></td><td>{esc(r['country'] or '')}</td><td class='muted'>{esc(c)}</td></tr>"
    body += "</table></div>"
    return page("Поставщики", body, "/suppliers")




def add_contact_link(sid):
    return f'<a href="/supplier?id={sid}" style="color:#d8a027">указать контакт</a>'


def contact_links(r):
    """Телефон как tel:, почта как mailto:, WeChat текстом."""
    parts = []
    if r["phone"]:
        tel = re.sub(r"[^\d+]", "", r["phone"])
        parts.append(f'<a href="tel:{tel}" style="color:#3d8bfd;text-decoration:none">📞 {esc(r["phone"])}</a>')
    if r["email"]:
        parts.append(f'<a href="mailto:{esc(r["email"])}" style="color:#3d8bfd;text-decoration:none">✉️ {esc(r["email"])}</a>')
    if r["wechat"]:
        parts.append(f'<span class="badge">WeChat: {esc(r["wechat"])}</span>')
    return " ".join(parts) if parts else ""


def view_supplier(sid):
    con = db()
    sup = con.execute("SELECT * FROM suppliers WHERE id=?", (sid,)).fetchone()
    if not sup:
        con.close()
        return page("Поставщик", "<div class='card'>Поставщик не найден. <a href='/suppliers'>Назад</a></div>")
    items = con.execute("""SELECT name, price, currency FROM items
                           WHERE supplier_id=? ORDER BY price LIMIT 500""", (sid,)).fetchall()
    cnt = con.execute("SELECT COUNT(*) c FROM items WHERE supplier_id=?", (sid,)).fetchone()["c"]
    con.close()

    tel = re.sub(r"[^\d+]", "", sup["phone"] or "")
    call = f'<a class="btn" href="tel:{tel}">Позвонить</a>' if tel else ""
    mail = f'<a class="btn gray" href="mailto:{esc(sup["email"])}">Написать</a>' if sup["email"] else ""

    body = f"""<div class="card">
<h2 style="margin:0 0 6px">{esc(sup["name"])}</h2>
<div class="muted">{esc(sup["country"] or "")} · позиций: {cnt}</div>
<div style="margin-top:10px">{contact_links(sup)}</div>
<div style="margin-top:12px">{call} {mail}</div>
</div>
<div class="card"><b>Контакты (можно изменить)</b>
<form method="post" action="/supplier/update">
<input type="hidden" name="id" value="{sid}">
<input type="text" name="name" value="{esc(sup["name"])}" placeholder="Компания">
<div class="row">
<input type="text" name="contact" value="{esc(sup["contact"] or "")}" placeholder="Контактное лицо">
<input type="text" name="phone" value="{esc(sup["phone"] or "")}" placeholder="Телефон / WhatsApp">
</div>
<div class="row">
<input type="text" name="email" value="{esc(sup["email"] or "")}" placeholder="E-mail">
<input type="text" name="wechat" value="{esc(sup["wechat"] or "")}" placeholder="WeChat">
<input type="text" name="country" value="{esc(sup["country"] or "")}" placeholder="Страна">
</div>
<input type="text" name="note" value="{esc(sup["note"] or "")}" placeholder="Примечание">
<button>Сохранить контакты</button>
</form></div>
<div class="card"><b>Товары поставщика</b><table>
<tr><th>Товар</th><th>Цена</th></tr>"""
    for it in items:
        body += f'<tr><td>{esc(it["name"])}</td><td class="p">{esc(it["price"])} {esc(it["currency"])}</td></tr>'
    body += "</table></div>"
    return page(sup["name"], body, "/suppliers")


def view_upload():
    con = db()
    sups = con.execute("SELECT id, name FROM suppliers ORDER BY name").fetchall()
    ups = con.execute("""SELECT u.*, s.name AS sup FROM uploads u
                         LEFT JOIN suppliers s ON s.id=u.supplier_id ORDER BY u.id DESC LIMIT 30""").fetchall()
    con.close()
    opts = "".join(f'<option value="{s["id"]}">{esc(s["name"])}</option>' for s in sups)
    body = f"""<div class="card"><form method="post" action="/upload/do" enctype="multipart/form-data">
<label class="muted">Поставщик</label>
<select name="supplier_id">{opts}</select>
<label class="muted">Файл прайса: xlsx, pdf, csv, txt</label>
<input type="file" name="file" accept=".xlsx,.xlsm,.pdf,.csv,.txt" required>
<button>Загрузить и разобрать</button>
<div class="muted">Приложение само найдёт колонки «наименование» и «цена».</div>
</form></div>"""
    body += '<div class="card"><table><tr><th>Файл</th><th>Поставщик</th><th>Позиций</th><th>Дата</th></tr>'
    for u in ups:
        body += f"<tr><td>{esc(u['filename'])}</td><td>{esc(u['sup'] or '—')}</td><td>{u['items']}</td><td class='muted'>{esc(u['created'])}</td></tr>"
    body += "</table></div>"
    return page("Прайс-листы", body, "/upload")


def view_export(q=""):
    body = """<div class="card"><p>Выгрузка всех позиций прайсов и списка поставщиков в Excel.</p>
<a class="btn" href="/export.xlsx">Скачать Excel</a></div>"""
    return page("Экспорт", body, "/export")


def make_excel():
    from openpyxl import Workbook
    from openpyxl.styles import Font
    con = db()
    wb = Workbook()
    ws = wb.active
    ws.title = "Позиции"
    ws.append(["Поставщик", "Страна", "Контакт", "Товар", "Цена", "Валюта", "Файл"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in con.execute("""SELECT s.name, s.country, s.contact, i.name, i.price, i.currency, u.filename
                            FROM items i JOIN suppliers s ON s.id=i.supplier_id
                            LEFT JOIN uploads u ON u.id=i.upload_id ORDER BY s.name, i.price"""):
        ws.append(list(r))
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 50
    ws2 = wb.create_sheet("Поставщики")
    ws2.append(["Компания", "Страна", "Контакт", "Телефон", "E-mail", "WeChat", "Примечание"])
    for c in ws2[1]:
        c.font = Font(bold=True)
    for r in con.execute("SELECT name,country,contact,phone,email,wechat,note FROM suppliers ORDER BY name"):
        ws2.append(list(r))
    con.close()
    path = os.path.join(APP_DIR, "export_all.xlsx")
    wb.save(path)
    return path



# ------------------------- история цен -------------------------
def _two_prices(supplier_id, name_lc):
    con = db()
    rows = con.execute("""SELECT price, upload_id FROM items
                          WHERE supplier_id=? AND name_lc=?
                          ORDER BY upload_id DESC, id DESC LIMIT 2""",
                       (supplier_id, name_lc)).fetchall()
    con.close()
    return rows


def price_changes():
    """Все товары, у которых цена изменилась между двумя последними прайсами поставщика."""
    con = db()
    seen = con.execute("""SELECT i.supplier_id, i.name_lc, i.name, s.name AS sup, s.country,
                                 s.contact, s.phone, s.wechat, s.email, i.currency
                          FROM items i JOIN suppliers s ON s.id=i.supplier_id
                          GROUP BY i.supplier_id, i.name_lc""").fetchall()
    con.close()
    out = []
    for r in seen:
        two = _two_prices(r["supplier_id"], r["name_lc"])
        if len(two) < 2:
            continue
        new, old = two[0]["price"], two[1]["price"]
        if not old or old == new:
            continue
        out.append({
            "name": r["name"], "supplier": r["sup"], "country": r["country"],
            "contact": " · ".join(x for x in [r["contact"], r["phone"], r["email"], r["wechat"]] if x),
            "phone": r["phone"] or "", "email": r["email"] or "", "wechat": r["wechat"] or "",
            "person": r["contact"] or "",
            "old": old, "new": new, "cur": r["currency"],
            "pct": (new - old) / old * 100.0,
        })
    out.sort(key=lambda x: abs(x["pct"]), reverse=True)
    return out


def last_change_for(supplier_id, name_lc):
    two = _two_prices(supplier_id, name_lc)
    if len(two) < 2:
        return None
    new, old = two[0]["price"], two[1]["price"]
    if not old or old == new:
        return None
    return {"old": old, "new": new, "pct": (new - old) / old * 100.0}


def stats_block():
    con = db()
    suppliers = con.execute("SELECT COUNT(*) c FROM suppliers").fetchone()["c"]
    uploads = con.execute("SELECT COUNT(*) c FROM uploads").fetchone()["c"]
    items = con.execute("SELECT COUNT(*) c FROM items").fetchone()["c"]
    con.close()
    ch = price_changes()
    up = [c for c in ch if c["pct"] > 0]
    down = [c for c in ch if c["pct"] < 0]
    return suppliers, uploads, items, up, down


def view_changes():
    ch = price_changes()
    sup, ups, items, up, down = stats_block()
    body = f"""<div class="card">
<div class="row"><div><div class="muted">Поставщиков</div><b>{sup}</b></div>
<div><div class="muted">Прайсов загружено</div><b>{ups}</b></div>
<div><div class="muted">Позиций</div><b>{items}</b></div>
<div><div class="muted">Подорожали</div><b style="color:#ff7b72">{len(up)}</b></div>
<div><div class="muted">Подешевели</div><b style="color:#2fbf71">{len(down)}</b></div></div></div>"""

    if not ch:
        body += ('<div class="card"><b>Изменений пока нет.</b>'
                 '<div class="muted">Загрузите второй прайс того же поставщика — '
                 'приложение сравнит цены и покажет, что подорожало или подешевело.</div></div>')
    else:
        body += '<div class="card"><table><tr><th>Товар</th><th>Было</th><th>Стало</th><th>Δ</th><th>Поставщик</th></tr>'
        for c in ch[:200]:
            color = "#ff7b72" if c["pct"] > 0 else "#2fbf71"
            sign = "+" if c["pct"] > 0 else ""
            body += (f'<tr><td>{esc(c["name"])}</td>'
                     f'<td class="muted">{esc(c["old"])} {esc(c["cur"])}</td>'
                     f'<td><b>{esc(c["new"])} {esc(c["cur"])}</b></td>'
                     f'<td style="color:{color};font-weight:700">{sign}{c["pct"]:.1f}%</td>'
                     f'<td>{esc(c["supplier"])}<div class="muted">{esc(c["contact"])}</div></td></tr>')
        body += "</table></div>"
    return page("Изменения цен", body, "/changes")



# ------------------------- заявка и КП -------------------------
def parse_order_items(path):
    """Из файла или фото достаём список «наименование — количество»."""
    ext = os.path.splitext(path)[1].lower()
    names = []
    if ext in (".png", ".jpg", ".jpeg", ".webp"):
        text = ocr_image_text(path)
        if text.strip():
            got = items_from_text(text)
            if got:
                return [(n, 1) for n, _p, _c, _u in got]
        try:
            import importlib.util as iu
            script = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                                  "..", "..", "Documents", "deepseek-harness",
                                                  "default-workspace", "shot_to_price.py"))
            spec = iu.spec_from_file_location("shot_to_price", script)
            mod = iu.module_from_spec(spec)
            spec.loader.exec_module(mod)
            for it in mod.parse_image(path, "заявка", write=False):
                names.append((it["name"], 1))
        except Exception as e:
            print("OCR error:", e)
        return names
    for name, price, cur, unit in parse_price_file(path):
        names.append((name, 1))
    return names


def match_item(name, limit=3):
    """Лучшие предложения из базы по названию из заявки."""
    rows, mode = smart_search(name, limit=limit)
    return rows[:limit], mode


def create_order(src_path, supplier_dummy=None):
    items = parse_order_items(src_path)
    con = db()
    cur = con.execute("INSERT INTO orders(created,filename,markup,total_buy,total_sell) VALUES(?,?,?,?,?)",
                      (datetime.now().strftime("%Y-%m-%d %H:%M"), os.path.basename(src_path), 20, 0, 0))
    oid = cur.lastrowid
    for name, qty in items:
        matches, _ = match_item(name)
        best = matches[0] if matches else None
        con.execute("""INSERT INTO order_items(order_id,req_name,qty,match_name,supplier_id,price,currency,usd)
                       VALUES(?,?,?,?,?,?,?,?)""",
                    (oid, name, qty,
                     best["name"] if best else "", best["sid"] if best else None,
                     best["price"] if best else None, best["cur"] if best else "",
                     best["usd"] if best else None))
    con.commit()
    con.close()
    return oid


def order_rows(oid):
    con = db()
    rows = con.execute("""SELECT oi.*, s.name AS sup, s.phone, s.email, s.wechat, s.contact
                          FROM order_items oi LEFT JOIN suppliers s ON s.id=oi.supplier_id
                          WHERE oi.order_id=? ORDER BY oi.id""", (oid,)).fetchall()
    order = con.execute("SELECT * FROM orders WHERE id=?", (oid,)).fetchone()
    con.close()
    return order, rows


def make_kp(oid, markup):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    order, rows = order_rows(oid)
    wb = Workbook()
    ws = wb.active
    ws.title = "КП"
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    st = get_settings()
    ws["A1"] = st["company_name"]
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = " · ".join(x for x in [st["company_inn"] and f"ИНН {st['company_inn']}",
                                      st["company_phone"], st["company_email"],
                                      st["company_address"]] if x)
    ws["A2"].font = Font(size=9, color="595959")
    ws["A3"] = f"КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ № {oid} от {datetime.now():%d.%m.%Y}   (наценка {markup:.0f}%)"
    ws["A3"].font = Font(bold=True, size=12)

    head = ["№", "Наименование в заявке", "Кол-во", "Предложение из базы", "Поставщик",
            "Контакт", "Закупка, USD", "Цена продажи, USD", "Сумма, USD"]
    for j, h in enumerate(head, start=1):
        c = ws.cell(row=5, column=j, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        c.border = border
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    total_buy = total_sell = 0.0
    r = 6
    for i, it in enumerate(rows, start=1):
        buy = it["usd"] or 0
        sell = round(buy * (1 + markup / 100.0), 2)
        qty = it["qty"] or 1
        contact = " · ".join(x for x in [it["contact"], it["phone"], it["email"], it["wechat"]] if x)
        vals = [i, it["req_name"], qty, it["match_name"] or "—", it["sup"] or "—", contact,
                buy or "—", sell or "—", round(sell * qty, 2)]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.border = border
            c.alignment = Alignment(vertical="top", wrap_text=(j in (2, 4, 6)))
        total_buy += buy * qty
        total_sell += sell * qty
        r += 1

    ws.cell(row=r, column=8, value="ИТОГО:").font = Font(bold=True)
    c = ws.cell(row=r, column=9, value=round(total_sell, 2))
    c.font = Font(bold=True)
    c.fill = PatternFill("solid", fgColor="FFF2CC")
    r += 1
    ws.cell(row=r, column=8, value="Закупка:").font = Font(bold=True)
    ws.cell(row=r, column=9, value=round(total_buy, 2))
    r += 1
    ws.cell(row=r, column=8, value="Прибыль:").font = Font(bold=True)
    ws.cell(row=r, column=9, value=round(total_sell - total_buy, 2))

    ws.cell(row=r + 1, column=1, value="Условия:").font = Font(bold=True, size=9)
    ws.cell(row=r + 2, column=1, value=st["kp_terms"]).font = Font(size=9)
    for j, w in enumerate([5, 34, 8, 40, 20, 26, 14, 16, 14], start=1):
        ws.column_dimensions[chr(64 + j)].width = w

    con = db()
    con.execute("UPDATE orders SET markup=?, total_buy=?, total_sell=? WHERE id=?",
                (markup, round(total_buy, 2), round(total_sell, 2), oid))
    con.commit()
    con.close()

    out = os.path.join(APP_DIR, f"kp_{oid}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    wb.save(out)
    cur = con = None
    con = db()
    cur = con.execute("INSERT INTO kp_history(order_id,created,markup,total_buy,total_sell,path) "
                      "VALUES(?,?,?,?,?,?)",
                      (oid, datetime.now().strftime("%Y-%m-%d %H:%M"), markup,
                       round(total_buy, 2), round(total_sell, 2), out))
    kp_id = cur.lastrowid
    con.commit()
    con.close()
    return out, kp_id


def view_order(oid=None):
    body = ""
    if oid:
        order, rows = order_rows(oid)
        if not order:
            return page("Заявка", "<div class='card'>Заявка не найдена. <a href='/order'>Новая</a></div>")
        total = sum((it["usd"] or 0) for it in rows)
        body += f"""<div class="card"><b>Заявка №{oid}</b>
<div class="muted">файл: {esc(order['filename'])} · позиций: {len(rows)} · закупка: {total:,.2f} USD</div></div>
<div class="card"><table>
<tr><th>#</th><th>В заявке</th><th>Предложение</th><th>Поставщик</th><th>Закупка</th><th>Контакт</th></tr>"""
        for i, it in enumerate(rows, start=1):
            contact = " · ".join(x for x in [it["contact"], it["phone"], it["email"], it["wechat"]] if x)
            body += (f"<tr><td>{i}</td><td>{esc(it['req_name'][:70])}</td>"
                     f"<td>{esc(it['match_name'] or '—')}</td><td>{esc(it['sup'] or '—')}</td>"
                     f"<td class='p'>{('%.2f' % it['usd']) + ' USD' if it['usd'] else '—'}</td>"
                     f"<td class='muted'>{esc(contact)}</td></tr>")
        body += "</table></div>"
        body += f"""<div class="card"><b>Следующий шаг: сделать КП?</b>
<form method="post" action="/order/kp">
<input type="hidden" name="id" value="{oid}">
<label class="muted">Какую наценку поставить, %</label>
<input type="text" name="markup" value="20">
<button>Сформировать КП</button>
</form>
<div class="muted">Считается цена продажи = закупка + наценка, итог и прибыль.</div></div>"""
        return page(f"Заявка №{oid}", body, "/order")

    # новая заявка
    body = """<div class="card"><b>Новая заявка</b>
<form method="post" action="/order/upload" enctype="multipart/form-data">
<label class="muted">Фото или файл заявки (png, jpg, xlsx, pdf, csv, txt)</label>
<input type="file" name="file" accept=".png,.jpg,.jpeg,.webp,.xlsx,.xlsm,.pdf,.csv,.txt" required>
<button>Проверить заявку по базе</button>
</form>
<div class="muted">Приложение распознаёт позиции, ищет их в базе поставщиков,
показывает закупочную цену и предложение.</div></div>"""
    con = db()
    orders = con.execute("SELECT id, created, filename, total_buy FROM orders ORDER BY id DESC LIMIT 20").fetchall()
    con.close()
    if orders:
        body += '<div class="card"><b>Последние заявки</b><table><tr><th>№</th><th>Дата</th><th>Файл</th><th>Закупка</th></tr>'
        for o in orders:
            body += (f'<tr><td><a href="/order?id={o["id"]}">{o["id"]}</a></td>'
                     f'<td class="muted">{esc(o["created"])}</td><td>{esc(o["filename"])}</td>'
                     f'<td class="p">{("%.2f" % o["total_buy"]) if o["total_buy"] else "—"}</td></tr>')
        body += "</table></div>"
    return page("Заявка", body, "/order")



# ------------------------- реквизиты компании -------------------------
DEFAULT_SETTINGS = {
    "company_name": "Ваша компания",
    "company_inn": "",
    "company_phone": "",
    "company_email": "",
    "company_address": "",
    "kp_terms": "Цены действительны 14 дней. Срок поставки уточняется при заказе. Гарантия — по условиям производителя.",
}


def get_settings():
    con = db()
    out = dict(DEFAULT_SETTINGS)
    for r in con.execute("SELECT key, value FROM settings"):
        out[r["key"]] = r["value"] or ""
    con.close()
    return out


def save_settings(form):
    con = db()
    for k in DEFAULT_SETTINGS:
        if k in form:
            con.execute("INSERT INTO settings(key,value) VALUES(?,?) "
                        "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (k, form[k]))
    con.commit()
    con.close()


def view_settings():
    st = get_settings()
    body = """<div class="card"><b>Реквизиты для КП</b>
<form method="post" action="/settings/save">
<label class="muted">Название компании</label>
<input type="text" name="company_name" value="{cn}">
<div class="row">
<input type="text" name="company_inn" value="{inn}" placeholder="ИНН">
<input type="text" name="company_phone" value="{phone}" placeholder="Телефон">
<input type="text" name="company_email" value="{email}" placeholder="E-mail">
</div>
<input type="text" name="company_address" value="{addr}" placeholder="Адрес">
<label class="muted">Условия в КП</label>
<textarea name="kp_terms" rows="3">{terms}</textarea>
<button>Сохранить реквизиты</button>
</form></div>""".format(cn=esc(st["company_name"]), inn=esc(st["company_inn"]),
                          phone=esc(st["company_phone"]), email=esc(st["company_email"]),
                          addr=esc(st["company_address"]), terms=esc(st["kp_terms"]))
    return page("Реквизиты", body, "/settings")


def view_kp_history():
    con = db()
    rows = con.execute("SELECT * FROM kp_history ORDER BY id DESC LIMIT 50").fetchall()
    con.close()
    body = '<div class="card"><b>История КП</b><table><tr><th>№</th><th>Дата</th><th>Заявка</th><th>Наценка</th><th>Закупка</th><th>Продажа</th><th></th></tr>'
    if not rows:
        body += '<tr><td colspan="7" class="muted">Пока нет КП — сделай первую из заявки.</td></tr>'
    for r in rows:
        body += (f'<tr><td>{r["id"]}</td><td class="muted">{esc(r["created"])}</td>'
                 f'<td><a href="/order?id={r["order_id"]}">{r["order_id"]}</a></td>'
                 f'<td>{r["markup"]:.0f}%</td><td>{r["total_buy"]:.2f}</td>'
                 f'<td class="p">{r["total_sell"]:.2f}</td>'
                 f'<td><a href="/kp/{r["id"]}.xlsx" class="btn gray" style="padding:6px 10px">скачать</a></td></tr>')
    body += "</table></div>"
    return page("История КП", body, "/kp")



# ------------------------- стартовая страница -------------------------
START_CSS = """
*{box-sizing:border-box;-webkit-text-size-adjust:100%}
body{margin:0;font:16px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;
background:#0b1016;color:#e8eef6}
.wrap{max-width:720px;margin:0 auto;padding:28px 18px 60px}
h1{font-size:26px;margin:6px 0 4px}
.sub{color:#93a3b8;margin-bottom:22px}
.card{background:#18202b;border:1px solid #2a3646;border-radius:16px;padding:20px;margin-bottom:16px}
.big{display:block;text-align:center;background:#3d8bfd;color:#fff;text-decoration:none;
padding:18px;border-radius:14px;font-size:18px;font-weight:700;margin-bottom:12px}
.big.green{background:#2fbf71;color:#08130c}
.big.gray{background:#243244}
.step{display:flex;gap:12px;margin:12px 0}
.num{background:#22303f;border-radius:10px;width:34px;height:34px;flex:0 0 34px;display:flex;
align-items:center;justify-content:center;font-weight:700;color:#3d8bfd}
.code{background:#0e141b;border:1px solid #2a3646;border-radius:8px;padding:8px 10px;
font-family:Consolas,monospace;font-size:13px;word-break:break-all;color:#9cc4ff}
.muted{color:#93a3b8;font-size:14px}
h2{font-size:18px;margin:0 0 10px}
"""


def view_start():
    url = ""
    try:
        with open(os.path.join(APP_DIR, "public_url.txt"), encoding="utf-8") as f:
            url = f.read().strip()
    except OSError:
        url = ""
    phone_block = ""
    if url:
        phone_block = f"""<div class="card"><h2>Телефон / планшет</h2>
<div class="muted">Открой эту ссылку в браузере телефона:</div>
<div class="code" style="margin:10px 0">{esc(url)}</div>
<a class="big" href="{esc(url)}">Открыть на телефоне</a>
<div class="muted">Safari → «Поделиться» → «На экран Домой» — будет иконка.</div></div>"""
    else:
        phone_block = """<div class="card"><h2>Телефон / планшет</h2>
<div class="muted">Ссылка появится после запуска (файл public_url.txt в папке приложения).</div></div>"""

    body = f"""<div class="wrap">
<h1>SupplierHub</h1>
<div class="sub">База поставщиков, прайсов, заявок и коммерческих предложений</div>

<div class="card"><h2>Уже установлено — открыть</h2>
<a class="big green" href="/">Открыть приложение (на этом устройстве)</a>
<div class="muted">Поиск товаров, заявка, КП, поставщики.</div></div>

{phone_block}

<div class="card"><h2>Windows-программа</h2>
<a class="big" href="/download">Скачать для Windows (архив)</a>
<div class="muted">Распакуй → двойной клик «Установить.bat» → готово.
Ярлык на рабочем столе, автозапуск при включении компьютера.</div></div>

<div class="card"><h2>Как пользоваться</h2>
<div class="step"><div class="num">1</div><div><b>Заявка</b><br>
<span class="muted">Загрузи фото, PDF или Excel заявки — приложение найдёт позиции в базе,
покажет закупочную цену и поставщика.</span></div></div>
<div class="step"><div class="num">2</div><div><b>Наценка</b><br>
<span class="muted">Укажи процент — система посчитает цену продажи и прибыль.</span></div></div>
<div class="step"><div class="num">3</div><div><b>КП</b><br>
<span class="muted">Нажми «Сформировать КП» — скачается готовое предложение
с реквизитами твоей компании.</span></div></div>
<div class="step"><div class="num">4</div><div><b>Поиск</b><br>
<span class="muted">Ищи товар по-русски или английски: дешёвое сверху зелёным,
рядом контакт поставщика и кнопка звонка.</span></div></div>
</div>

<div class="card"><h2>Удаление</h2>
<div class="muted">В папке программы запусти «Удалить.bat» — удалит ярлыки, автозагрузку и базу.</div></div>
</div>"""
    return f"""<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SupplierHub — установка и запуск</title><style>{START_CSS}</style></head>
<body>{body}</body></html>"""


# ------------------------- сервер -------------------------
class Handler(BaseHTTPRequestHandler):
    server_version = "SupplierHub/1.0"

    def _send(self, code, body, ctype="text/html; charset=utf-8", extra=None):
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        qs = urllib.parse.parse_qs(parsed.query)
        q = (qs.get("q") or [""])[0]

        if path == "/":
            return self._send(200, view_search(q))
        if path == "/suppliers":
            return self._send(200, view_suppliers())
        if path == "/upload":
            return self._send(200, view_upload())
        if path == "/supplier":
            try:
                sid = int((qs.get("id") or ["0"])[0])
            except ValueError:
                sid = 0
            return self._send(200, view_supplier(sid))
        if path == "/order":
            oid = None
            if qs.get("id"):
                try:
                    oid = int(qs["id"][0])
                except ValueError:
                    oid = None
            return self._send(200, view_order(oid))
        if path == "/start":
            return self._send(200, view_start())
        if path == "/download":
            zip_path = os.path.join(APP_DIR, "SupplierHub_Windows.zip")
            if not os.path.exists(zip_path):
                alt = os.path.join(os.path.expanduser("~"), "Desktop", "SupplierHub_Windows.zip")
                zip_path = alt if os.path.exists(alt) else zip_path
            if os.path.exists(zip_path):
                with open(zip_path, "rb") as f:
                    data = f.read()
                return self._send(200, data, "application/zip",
                                  {"Content-Disposition": 'attachment; filename="SupplierHub_Windows.zip"'})
            return self._send(404, page("Скачать", "<div class='card'>Архив не найден.</div>"))
        if path == "/settings":
            return self._send(200, view_settings())
        if path == "/kp":
            return self._send(200, view_kp_history())
        if path.startswith("/kp/") and path.endswith(".xlsx"):
            kid = re.sub(r"\D", "", path)
            con = db()
            row = con.execute("SELECT path FROM kp_history WHERE id=?", (int(kid or 0),)).fetchone()
            con.close()
            if row and os.path.exists(row["path"]):
                with open(row["path"], "rb") as f:
                    data = f.read()
                return self._send(200, data,
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  {"Content-Disposition": f'attachment; filename="KP_{kid}.xlsx"'})
            return self._send(404, page("КП", "<div class='card'>Файл не найден</div>"))
        if path == "/changes":
            return self._send(200, view_changes())
        if path == "/export":
            return self._send(200, view_export())
        if path == "/manifest.json":
            manifest = {
                "name": "SupplierHub", "short_name": "SupplierHub", "start_url": "/",
                "display": "standalone", "background_color": "#10151c", "theme_color": "#10151c",
                "icons": [{"src": "/icon.png", "sizes": "192x192", "type": "image/png"}],
            }
            return self._send(200, json.dumps(manifest), "application/manifest+json")
        if path == "/icon.png":
            return self._icon()
        if path == "/export.xlsx":
            p = make_excel()
            with open(p, "rb") as f:
                data = f.read()
            return self._send(200, data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              {"Content-Disposition": 'attachment; filename="suppliers.xlsx"'})
        return self._send(404, page("404", "<div class='card'>Страница не найдена. <a href='/'>На главную</a></div>"))

    def _icon(self):
        # простой PNG-значок 192×192, генерируем один раз
        p = os.path.join(APP_DIR, "icon.png")
        if not os.path.exists(p):
            try:
                from PIL import Image, ImageDraw
                im = Image.new("RGB", (192, 192), (16, 21, 28))
                d = ImageDraw.Draw(im)
                d.rectangle([40, 88, 152, 104], fill=(61, 139, 253))
                d.rectangle([40, 120, 152, 136], fill=(61, 139, 253))
                d.rectangle([40, 56, 152, 72], fill=(61, 139, 253))
                im.save(p)
            except Exception:
                return self._send(404, b"", "image/png")
        with open(p, "rb") as f:
            return self._send(200, f.read(), "image/png")

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        ctype = self.headers.get("Content-Type", "")
        body = self.rfile.read(length)

        if self.path == "/order/upload":
            fields, files = parse_multipart(body, ctype)
            f = files.get("file")
            if not f:
                return self._send(400, page("Заявка", "<div class='card'>Файл не получен</div>"))
            fname, data = f
            safe = re.sub(r"[^\w\.\-]+", "_", os.path.basename(fname))[:120]
            dest = os.path.join(UP_DIR, f"order_{datetime.now():%Y%m%d_%H%M%S}_{safe}")
            with open(dest, "wb") as fh:
                fh.write(data)
            oid = create_order(dest)
            return self._redirect(f"/order?id={oid}")

        if self.path == "/order/kp":
            if "multipart/form-data" in ctype:
                form = parse_multipart(body, ctype)[0]
                form = {k: [v] for k, v in form.items()}
            else:
                form = urllib.parse.parse_qs(body.decode("utf-8", "ignore"))
            oid = int((form.get("id") or ["0"])[0] or 0)
            try:
                markup = float((form.get("markup") or ["20"])[0])
            except ValueError:
                markup = 20.0
            path_xlsx, kp_id = make_kp(oid, markup)
            with open(path_xlsx, "rb") as fh:
                data = fh.read()
            return self._send(200, data,
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              {"Content-Disposition": f'attachment; filename="KP_{oid}.xlsx"'})

        if self.path == "/settings/save":
            form = urllib.parse.parse_qs(body.decode("utf-8", "ignore"))
            save_settings({k: (v[0] if v else "") for k, v in form.items()})
            return self._redirect("/settings")

        if self.path == "/supplier/update":
            form = urllib.parse.parse_qs(body.decode("utf-8", "ignore"))
            g = lambda k: (form.get(k) or [""])[0].strip()
            con = db()
            con.execute("""UPDATE suppliers SET name=?, contact=?, phone=?, email=?, wechat=?,
                           country=?, note=? WHERE id=?""",
                        (g("name"), g("contact"), g("phone"), g("email"), g("wechat"),
                         g("country"), g("note"), int(g("id") or 0)))
            con.commit()
            con.close()
            return self._redirect(f"/supplier?id={g('id')}")

        if self.path == "/suppliers/add":
            form = urllib.parse.parse_qs(body.decode("utf-8", "ignore"))
            get = lambda k: (form.get(k) or [""])[0].strip()
            con = db()
            con.execute("""INSERT INTO suppliers(created,name,country,contact,phone,email,wechat,note)
                           VALUES(?,?,?,?,?,?,?,?)""",
                        (datetime.now().strftime("%Y-%m-%d"), get("name"), get("country"), get("contact"),
                         get("phone"), get("email"), get("wechat"), get("note")))
            con.commit()
            con.close()
            return self._redirect("/suppliers")

        if self.path == "/upload/do":
            fields, files = parse_multipart(body, ctype)
            supplier_id = int(fields.get("supplier_id", "0") or 0)
            f = files.get("file")
            if not f:
                return self._send(400, page("Ошибка", "<div class='card'>Файл не получен</div>"))
            fname, data = f
            safe = re.sub(r"[^\w\.\-]+", "_", os.path.basename(fname))[:120]
            dest = os.path.join(UP_DIR, f"{datetime.now():%Y%m%d_%H%M%S}_{safe}")
            with open(dest, "wb") as fh:
                fh.write(data)
            import hashlib
            fhash = hashlib.sha1(data).hexdigest()
            con = db()
            dup = con.execute("SELECT id FROM uploads WHERE hash=?", (fhash,)).fetchone()
            if dup:
                con.close()
                try:
                    os.remove(dest)
                except OSError:
                    pass
                return self._redirect("/upload?dup=1")
            ext_low = os.path.splitext(dest)[1].lower()
            if ext_low in (".png", ".jpg", ".jpeg", ".webp"):
                items = parse_screenshot(dest)
            else:
                items = parse_price_file(dest)
            cur = con.execute("INSERT INTO uploads(created,supplier_id,filename,items,hash) VALUES(?,?,?,?,?)",
                              (datetime.now().strftime("%Y-%m-%d %H:%M"), supplier_id, safe, len(items), fhash))
            up_id = cur.lastrowid
            seen = set()
            uniq = []
            for it in items:
                k = (normalize(it[0]), round(it[1], 2) if it[1] else 0)
                if k in seen:
                    continue
                seen.add(k)
                uniq.append(it)
            items = uniq
            con.execute("UPDATE uploads SET items=? WHERE id=?", (len(items), up_id))
            invalidate_cache()
            for name, price, cur_code, unit in items:
                con.execute("INSERT INTO items(upload_id,supplier_id,name,name_lc,price,currency,unit,raw) VALUES(?,?,?,?,?,?,?,?)",
                            (up_id, supplier_id, name, name.lower(), price, cur_code, unit, ""))
            con.commit()
            con.close()
            return self._redirect(f"/upload")

        return self._send(404, "not found")

    def _redirect(self, to):
        self.send_response(303)
        self.send_header("Location", to)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def log_message(self, fmt, *args):
        pass


def parse_multipart(body, ctype):
    m = re.search(r"boundary=([^;]+)", ctype)
    fields, files = {}, {}
    if not m:
        return fields, files
    boundary = m.group(1).strip().strip('"').encode()
    for part in body.split(b"--" + boundary):
        if b"Content-Disposition" not in part:
            continue
        head, _, data = part.partition(b"\r\n\r\n")
        head_txt = head.decode("utf-8", "ignore")
        nm = re.search(r'name="([^"]+)"', head_txt)
        fn = re.search(r'filename="([^"]*)"', head_txt)
        if not nm:
            continue
        if fn and fn.group(1):
            files[nm.group(1)] = (fn.group(1), data.rstrip(b"\r\n-"))
        else:
            fields[nm.group(1)] = data.decode("utf-8", "ignore").strip()
    return fields, files


def local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        con = db()
        con.execute("INSERT INTO suppliers(created,name,country,contact,phone,email,wechat,note) VALUES(?,?,?,?,?,?,?,?)",
                    (datetime.now().strftime("%Y-%m-%d"), "Test Trading Co", "Китай", "Li Wei", "+86 138 0000", "li@test.cn", "liwei88", "тест"))
        con.commit()
        print("suppliers:", con.execute("SELECT COUNT(*) FROM suppliers").fetchone()[0])
        con.close()
        print("SELFTEST OK")
    else:
        srv = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
        ip = local_ip()
        print("=" * 60)
        print("SupplierHub запущен")
        print(f"  Компьютер:  http://localhost:{PORT}")
        print(f"  iPhone:     http://{ip}:{PORT}   (в той же Wi-Fi сети)")
        print("  Остановить: Ctrl+C")
        print("=" * 60)
        srv.serve_forever()
