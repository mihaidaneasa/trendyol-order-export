#!/usr/bin/env python3
import os, pathlib, requests, json, sys, base64, time, tempfile
import pandas as pd
from dotenv import load_dotenv
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime, timedelta, timezone

# ── Base paths ────────────────────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    CODE_BASE = pathlib.Path(sys._MEIPASS)
    EXE_DIR   = pathlib.Path(sys.executable).resolve().parent
else:
    CODE_BASE = pathlib.Path(__file__).resolve().parent
    EXE_DIR   = CODE_BASE

# ── Baze de căi ─────────────────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    EXE_DIR    = pathlib.Path(sys.executable).resolve().parent      # unde e EXE-ul
    BUNDLE_DIR = pathlib.Path(getattr(sys, '_MEIPASS', EXE_DIR))    # unde PyInstaller dezarhivează fișierele
else:
    EXE_DIR    = pathlib.Path(__file__).resolve().parent
    BUNDLE_DIR = EXE_DIR

# ───────────────────────────────────────────────────────────────
# Încărcare configurație Trendyol DOAR din env.dpapi (fără .env)
# ───────────────────────────────────────────────────────────────

from encrypt_env_dpapi import dpapi_unprotect

# Determinăm folderul în care se află scriptul/exe-ul
if getattr(sys, 'frozen', False):
    BASE = pathlib.Path(sys.executable).resolve().parent
else:
    BASE = pathlib.Path(__file__).resolve().parent

def load_env_from_dpapi():
    """
    Caută env.dpapi în:
      1. lângă EXE (EXE_DIR)
      2. în folderul bundle (BUNDLE_DIR = _MEIPASS la onefile)
    și decriptează configurarea Trendyol.
    """
    enc_path = None
    for base in (EXE_DIR, BUNDLE_DIR):
        cand = base / "env.dpapi"
        if cand.exists():
            enc_path = cand
            break

    if not enc_path:
        raise RuntimeError(
            "Fișierul criptat env.dpapi nu a fost găsit în:\n"
            f" - {EXE_DIR}\n"
            f" - {BUNDLE_DIR}\n"
            "Fără acest fișier, nu pot încărca TRENDYOL_ID, KEY, SECRET."
        )

    dec_bytes = dpapi_unprotect(enc_path.read_bytes())

    with tempfile.NamedTemporaryFile("wb", delete=False) as tmp:
        tmp.write(dec_bytes)
        tmp.flush()
        load_dotenv(tmp.name, override=True)

    print(f"🔐 Configurația Trendyol a fost încărcată din: {enc_path}")

load_env_from_dpapi()

TRENDYOL_ID     = os.getenv("TRENDYOL_ID")
TRENDYOL_KEY    = os.getenv("TRENDYOL_KEY")
TRENDYOL_SECRET = os.getenv("TRENDYOL_SECRET")
TRENDYOL_REF    = os.getenv("TRENDYOL_REF", "TrendyolOblio/1.0")

if not all([TRENDYOL_ID, TRENDYOL_KEY, TRENDYOL_SECRET]):
    raise RuntimeError(
        "Config incompletă în env.dpapi. Verifică dacă sunt setate:\n"
        " - TRENDYOL_ID\n"
        " - TRENDYOL_KEY\n"
        " - TRENDYOL_SECRET"
    )

BASE_URL   = f"https://api.trendyol.com/sapigw/suppliers/{TRENDYOL_ID}"
ORDERS_URL = f"{BASE_URL}/orders"

# ── Auth headers ──────────────────────────────────────────────────────────────
TOKEN = base64.b64encode(f"{TRENDYOL_KEY}:{TRENDYOL_SECRET}".encode()).decode()
HEADERS = {
    "User-Agent": TRENDYOL_REF,
    "Authorization": f"Basic {TOKEN}",
    "Accept": "application/json",
    "Content-Type": "application/json"
}

# ── Coduri poștale ────────────────────────────────────────────────────────────
POSTAL_FILE = None
for p in (EXE_DIR / "coduri_postale_judete.json", CODE_BASE / "coduri_postale_judete.json"):
    if p.exists():
        POSTAL_FILE = p
        break

POSTAL_MAP = {}
if POSTAL_FILE:
    try:
        with open(POSTAL_FILE, "r", encoding="utf-8") as f:
            postal_data = json.load(f)
        POSTAL_MAP = {
            str(item.get("cod_postal", "")).zfill(6): {
                "judet": (item.get("judet") or "").strip(),
                "localitate": (item.get("localitate") or item.get("oras") or "").strip()
            }
            for item in postal_data if str(item.get("cod_postal", "")).strip()
        }
    except Exception as e:
        print(f"⚠️ Eroare la citirea {POSTAL_FILE.name}: {e}")
else:
    print("⚠️ ATENȚIE: coduri_postale_judete.json nu a fost găsit.")

# ── Helpers ───────────────────────────────────────────────────────────────────
def _first(*vals) -> str:
    for v in vals:
        if v is None:
            continue
        s = str(v).strip()
        if s:
            return s
    return ""

def _normalize_name(name: str) -> str:
    if not name:
        return ""
    name = " ".join(name.split())
    return " ".join([w.capitalize() for w in name.split()])

def _name_from(addr: dict, order: dict) -> str:
    raw = _first(
        (addr or {}).get("fullName"),
        f"{_first((addr or {}).get('firstName'))} {_first((addr or {}).get('lastName'))}".strip(),
        (order.get("buyer") or {}).get("name"),
        (order.get("customer") or {}).get("name")
    )
    return _normalize_name(raw)

def _city(addr: dict) -> str:
    return _first((addr or {}).get("district"), (addr or {}).get("town"), (addr or {}).get("city"))

def _postal(addr: dict) -> str:
    cp = _first((addr or {}).get("zipCode"), (addr or {}).get("postalCode"))
    return str(cp).zfill(6) if cp else ""

def _county_from_postal(cp: str) -> str:
    if not cp:
        return ""
    return (POSTAL_MAP.get(str(cp).zfill(6), {}) or {}).get("judet", "")

def _city_from_postal(cp: str) -> str:
    if not cp:
        return ""
    return (POSTAL_MAP.get(str(cp).zfill(6), {}) or {}).get("localitate", "")

def _county(addr: dict) -> str:
    cp = _postal(addr)
    from_cp = _county_from_postal(cp)
    if from_cp:
        return from_cp
    return _first((addr or {}).get("province"), (addr or {}).get("state"), (addr or {}).get("city"))

def _addr_clean(addr: dict) -> str:
    a1 = _first((addr or {}).get("address1"))
    a2 = _first((addr or {}).get("address2"))
    full = _first((addr or {}).get("fullAddress"))
    raw = _first((addr or {}).get("address"))
    parts = [p for p in (a1, a2) if p]
    if parts:
        return " ".join(parts).strip()
    if full:
        return full
    if raw:
        return raw
    pieces = [_first((addr or {}).get(k)) for k in ("street", "apartment", "block", "building")]
    pieces = [p for p in pieces if p]
    return " ".join(pieces).strip()

def _addr_tuple(prefix: str, addr: dict) -> dict:
    cp = _postal(addr)
    return {
        f"Adresa_{prefix}": _addr_clean(addr),
        f"Localitate_{prefix}": _city(addr),
        f"LocalitateCod_{prefix}": _city_from_postal(cp),
        f"Judet_{prefix}": _county(addr),
        f"CodPostal_{prefix}": cp,
    }

# ── API helpers ───────────────────────────────────────────────────────────────
from datetime import datetime, timedelta

def fetch_uninvoiced_order_numbers(
    statuses=("Created","Picking","Invoiced","Shipped","Delivered","UnDelivered","Returned","Repack","AtCollectionPoint")
) -> list[str]:
    results = set()

    now = datetime.now(timezone.utc)  # fix warning + UTC corect
    start_limit = now - timedelta(days=60)

    window_end = now
    while window_end > start_limit:
        window_start = max(window_end - timedelta(days=14), start_limit)

        start_ms = int(window_start.timestamp() * 1000)
        end_ms   = int(window_end.timestamp() * 1000)

        for status in statuses:
            page = 0
            while True:
                r = requests.get(
                    ORDERS_URL,
                    params={
                        "status": status,
                        "startDate": start_ms,
                        "endDate": end_ms,
                        "page": page,
                        "size": 100
                    },
                    headers=HEADERS,
                    timeout=25
                )
                r.raise_for_status()
                payload = r.json() or {}
                chunk = payload.get("content", []) or []

                if not chunk:
                    break

                for o in chunk:
                    # păstrează logica ta: "nefacturat" = fără invoiceLink
                    if not o.get("invoiceLink"):
                        results.add(str(_first(o.get("orderNumber"), o.get("id"))))

                if len(chunk) < 100:
                    break
                page += 1

        window_end = window_start
        time.sleep(0.1)

    return sorted(results)

def fetch_order_by_number(order_number: str) -> dict:
    """Citește comanda filtrând /orders după orderNumber (evită 403 de la /orders/{id})."""
    r = requests.get(
        ORDERS_URL,
        params={"orderNumber": order_number, "page": 0, "size": 1},
        headers=HEADERS,
        timeout=25
    )
    r.raise_for_status()
    payload = r.json() or {}
    content = payload.get("content") or []
    return content[0] if content else {}

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    try:
        order_numbers = fetch_uninvoiced_order_numbers()
    except requests.RequestException as e:
        tk.Tk().withdraw()
        messagebox.showerror("Eroare Trendyol", f"Nu am putut lista comenzile: {e}")
        return

    if not order_numbers:
        tk.Tk().withdraw()
        messagebox.showinfo("Rezultat", "ℹ️ Nu există comenzi fără factură.")
        return

    rows = []
    for onr in order_numbers:
        try:
            detail = fetch_order_by_number(onr)
        except requests.RequestException as e:
            print(f"⚠️ Order {onr}: eroare la citire prin filtrare: {e}")
            continue

        # Adrese: încercăm invoiceAddress, apoi billingAddress; shipping separat
        bill = (detail.get("invoiceAddress") or detail.get("billingAddress") or {})  # <—
        ship = (detail.get("shipmentAddress") or {})

        client_name = _name_from(bill, detail) or _name_from(ship, detail)

        row = {
            "OrderNumber": onr,
            "Client": client_name,
            "Status": detail.get("status", ""),
        }
        row.update(_addr_tuple("Facturare", bill))
        row.update(_addr_tuple("Livrare", ship))
        rows.append(row)

        # mică pauză, să nu lovim rate limits
        time.sleep(0.05)

    if not rows:
        tk.Tk().withdraw()
        messagebox.showinfo("Rezultat", "ℹ️ Nu am putut extrage adresele pentru comenzile găsite.")
        return

    columns = [
        "OrderNumber","Client","Status",
        "Adresa_Facturare","Localitate_Facturare","LocalitateCod_Facturare","Judet_Facturare","CodPostal_Facturare",
        "Adresa_Livrare","Localitate_Livrare","LocalitateCod_Livrare","Judet_Livrare","CodPostal_Livrare",
    ]
    df = pd.DataFrame(rows, columns=columns)

    root = tk.Tk(); root.withdraw()
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files","*.xlsx")],
        title="Salvează fișierul cu adrese (Facturare + Livrare)",
        initialfile="trendyol_clients.xlsx"
    )
    if not file_path:
        return

    df.to_excel(file_path, index=False)
    messagebox.showinfo("Succes", f"✅ Fișier salvat:\n{file_path}\n\nComenzi exportate: {len(rows)}")

if __name__ == "__main__":
    main()
