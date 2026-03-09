#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Utilitar DPAPI:

- dpapi_protect(data: bytes)  -> bytes  (criptează)
- dpapi_unprotect(data: bytes)-> bytes  (decriptează)

Dacă este rulat direct:
    python encrypt_env_dpapi.py .env env.dpapi

Criptează .env -> env.dpapi (legat de userul Windows curent) și poate șterge .env.
"""

import ctypes
import ctypes.wintypes as wt
from pathlib import Path
import sys

# --- Structuri și funcții DPAPI comune ---------------------------------------

class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wt.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_byte)),
    ]

def _blob_from_bytes(data: bytes) -> DATA_BLOB:
    buf = ctypes.create_string_buffer(data)
    return DATA_BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_byte)))

def dpapi_protect(data: bytes) -> bytes:
    """Criptează bytes folosind DPAPI (legat de userul Windows curent)."""
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32

    in_blob = _blob_from_bytes(data)
    out_blob = DATA_BLOB()

    if not crypt32.CryptProtectData(
        ctypes.byref(in_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(out_blob),
    ):
        raise ctypes.WinError()

    try:
        result = ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        kernel32.LocalFree(out_blob.pbData)

    return result

def dpapi_unprotect(data: bytes) -> bytes:
    """Decriptează bytes protejați cu DPAPI."""
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32

    in_blob = _blob_from_bytes(data)
    out_blob = DATA_BLOB()

    if not crypt32.CryptUnprotectData(
        ctypes.byref(in_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(out_blob),
    ):
        raise ctypes.WinError()

    try:
        result = ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        kernel32.LocalFree(out_blob.pbData)

    return result

# --- CLI simplu: python encrypt_env_dpapi.py input output --------------------

def _main_cli():
    if len(sys.argv) != 3:
        print("Usage: python encrypt_env_dpapi.py <input_env> <output_dpapi>")
        print("Ex:    python encrypt_env_dpapi.py .env env.dpapi")
        sys.exit(1)

    in_path = Path(sys.argv[1])
    out_path = Path(sys.argv[2])

    if not in_path.exists():
        raise FileNotFoundError(f"⚠️  Fișierul de intrare nu există: {in_path}")

    plain = in_path.read_bytes()
    cipher = dpapi_protect(plain)
    out_path.write_bytes(cipher)
    print(f"✅ Fișier criptat: {out_path} (legat de userul Windows curent)")

    # Ștergere opțională .env
    wipe = True
    if wipe:
        in_path.write_bytes(b"")
        in_path.unlink(missing_ok=True)
        print(f"🧹 Fișier {in_path} șters după criptare.")

if __name__ == "__main__":
    _main_cli()
