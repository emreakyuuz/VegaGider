"""
Vega Gider Raporlama - Yerel Masaüstü Sürümü (tamamen offline).

- Gömülü SQLite veritabanı kullanır (kurulum gerektirmez).
- React arayüzünü ve API'yi tek sunucudan sunar.
- Başlatıldığında tarayıcıyı otomatik açar.
- İnternet bağlantısı gerektirmez.
"""
import os
import sys
import json
import uuid
import sqlite3
import threading
import webbrowser
from pathlib import Path
from datetime import datetime, timezone
from urllib.parse import quote

import uvicorn
from fastapi import FastAPI, APIRouter, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from pydantic import BaseModel
from typing import Optional

import vega_processor as vp

PORT = int(os.environ.get("VEGA_PORT", "8777"))

# Çalışma dizini (PyInstaller onefile için _MEIPASS)
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys._MEIPASS)  # type: ignore
else:
    BASE_DIR = Path(__file__).parent
STATIC_DIR = BASE_DIR / "static"

# Veritabanı kullanıcının kişisel klasöründe kalıcı saklanır
DATA_DIR = Path.home() / "VegaGiderRaporlama"
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "veri.db"


# ----------------------------- SQLite -----------------------------
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.execute(
        """CREATE TABLE IF NOT EXISTS reports (
            id TEXT PRIMARY KEY,
            filename TEXT,
            period TEXT,
            created_at TEXT,
            transactions TEXT,
            vega_declared TEXT
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS mappings (
            code TEXT PRIMARY KEY,
            name TEXT,
            category TEXT,
            confirmed INTEGER
        )"""
    )
    conn.commit()
    # seed
    cur = conn.execute("SELECT COUNT(*) AS c FROM mappings")
    if cur.fetchone()["c"] == 0:
        conn.executemany(
            "INSERT INTO mappings(code,name,category,confirmed) VALUES(?,?,?,1)",
            [(c, n, cat) for c, n, cat in vp.SEED_MAPPINGS],
        )
        conn.commit()
    conn.close()


def mappings_dict():
    conn = get_db()
    rows = conn.execute("SELECT * FROM mappings").fetchall()
    conn.close()
    return {r["code"]: {"code": r["code"], "name": r["name"], "category": r["category"], "confirmed": bool(r["confirmed"])} for r in rows}


# ----------------------------- App -----------------------------
app = FastAPI(title="Vega Gider Raporlama")
api_router = APIRouter(prefix="/api")


class MappingItem(BaseModel):
    code: str
    name: Optional[str] = None
    category: str
    confirmed: bool = True


class TransactionUpdate(BaseModel):
    code: Optional[str] = None
    name: Optional[str] = None
    date: Optional[str] = None
    desc: Optional[str] = None
    borc: Optional[float] = None
    alacak: Optional[float] = None


class ResetRequest(BaseModel):
    clear_mappings: bool = False


def load_report_doc(report_id):
    conn = get_db()
    r = conn.execute("SELECT * FROM reports WHERE id=?", (report_id,)).fetchone()
    conn.close()
    if not r:
        return None
    return {
        "id": r["id"],
        "filename": r["filename"],
        "period": json.loads(r["period"] or "{}"),
        "created_at": r["created_at"],
        "transactions": json.loads(r["transactions"] or "[]"),
        "vega_declared": json.loads(r["vega_declared"] or "{}"),
    }


def assemble(doc):
    md = mappings_dict()
    groups, summary, charts = vp.compute_report(doc["transactions"], md)
    cat_by_code = {g["code"]: g["category"] for g in groups}
    txs = []
    for t in doc["transactions"]:
        tt = dict(t)
        tt["category"] = cat_by_code.get(t["code"], vp.default_category(t["code"]))
        txs.append(tt)
    validation = vp.build_validation(doc["transactions"], doc.get("vega_declared", {}))
    new_codes = []
    for g in groups:
        m = md.get(g["code"])
        if m is None or not m.get("confirmed", True):
            new_codes.append({"code": g["code"], "name": g["name"], "category": g["category"], "net": g["net"]})
    return {
        "id": doc["id"], "filename": doc.get("filename"), "period": doc.get("period", {}),
        "created_at": doc.get("created_at"), "transactions": txs, "groups": groups,
        "summary": summary, "charts": charts, "validation": validation, "new_codes": new_codes,
    }


def save_transactions(report_id, transactions):
    conn = get_db()
    conn.execute("UPDATE reports SET transactions=? WHERE id=?", (json.dumps(transactions, ensure_ascii=False), report_id))
    conn.commit()
    conn.close()


@api_router.get("/")
def root():
    return {"message": "Vega Gider Raporlama (Yerel)"}


@api_router.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    fname = file.filename or "veri.xlsx"
    if not (fname.lower().endswith(".xls") or fname.lower().endswith(".xlsx")):
        raise HTTPException(status_code=400, detail="Lütfen .xls veya .xlsx uzantılı bir dosya yükleyin.")
    content = await file.read()
    try:
        transactions, vega_declared = vp.read_excel(content, fname)
    except vp.VegaError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Dosya işlenirken bir hata oluştu: {e}")

    period = vp.detect_period(transactions)
    conn = get_db()
    md = mappings_dict()
    for t in transactions:
        code = t["code"]
        if code not in md:
            cat = vp.default_category(code)
            conn.execute("INSERT OR IGNORE INTO mappings(code,name,category,confirmed) VALUES(?,?,?,0)", (code, t["name"], cat))
            md[code] = {"code": code, "name": t["name"], "category": cat, "confirmed": False}
    rid = str(uuid.uuid4())
    doc = {
        "id": rid, "filename": fname, "period": period,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "transactions": transactions, "vega_declared": vega_declared,
    }
    conn.execute(
        "INSERT INTO reports(id,filename,period,created_at,transactions,vega_declared) VALUES(?,?,?,?,?,?)",
        (rid, fname, json.dumps(period, ensure_ascii=False), doc["created_at"],
         json.dumps(transactions, ensure_ascii=False), json.dumps(vega_declared, ensure_ascii=False)),
    )
    conn.commit()
    conn.close()
    return assemble(doc)


@api_router.get("/reports")
def list_reports():
    conn = get_db()
    rows = conn.execute("SELECT id,filename,period,created_at FROM reports").fetchall()
    conn.close()
    out = [{"id": r["id"], "filename": r["filename"], "period": json.loads(r["period"] or "{}"), "created_at": r["created_at"]} for r in rows]
    out.sort(key=lambda d: d.get("created_at", ""), reverse=True)
    return out


@api_router.get("/reports/{report_id}")
def get_report(report_id: str):
    doc = load_report_doc(report_id)
    if not doc:
        raise HTTPException(status_code=404, detail="Rapor bulunamadı.")
    return assemble(doc)


@api_router.delete("/reports/{report_id}")
def delete_report(report_id: str):
    conn = get_db()
    cur = conn.execute("DELETE FROM reports WHERE id=?", (report_id,))
    conn.commit()
    deleted = cur.rowcount
    conn.close()
    if deleted == 0:
        raise HTTPException(status_code=404, detail="Rapor bulunamadı.")
    return {"ok": True}


@api_router.put("/reports/{report_id}/transaction/{tx_id}")
def update_transaction(report_id: str, tx_id: str, upd: TransactionUpdate):
    doc = load_report_doc(report_id)
    if not doc:
        raise HTTPException(status_code=404, detail="Rapor bulunamadı.")
    found = False
    for t in doc["transactions"]:
        if t["id"] == tx_id:
            data = upd.model_dump(exclude_none=True)
            for k, v in data.items():
                t[k] = round(float(v), 2) if k in ("borc", "alacak") else v
            found = True
            break
    if not found:
        raise HTTPException(status_code=404, detail="Hareket bulunamadı.")
    save_transactions(report_id, doc["transactions"])
    return assemble(doc)


@api_router.delete("/reports/{report_id}/transaction/{tx_id}")
def delete_transaction(report_id: str, tx_id: str):
    doc = load_report_doc(report_id)
    if not doc:
        raise HTTPException(status_code=404, detail="Rapor bulunamadı.")
    before = len(doc["transactions"])
    doc["transactions"] = [t for t in doc["transactions"] if t["id"] != tx_id]
    if len(doc["transactions"]) == before:
        raise HTTPException(status_code=404, detail="Hareket bulunamadı.")
    save_transactions(report_id, doc["transactions"])
    return assemble(doc)


@api_router.get("/mappings")
def get_mappings():
    conn = get_db()
    rows = conn.execute("SELECT * FROM mappings").fetchall()
    conn.close()
    out = [{"code": r["code"], "name": r["name"], "category": r["category"], "confirmed": bool(r["confirmed"])} for r in rows]
    out.sort(key=lambda d: d["code"])
    return out


@api_router.put("/mappings")
def upsert_mapping(item: MappingItem):
    if item.category not in ("DOKUMA", "BOYAHANE", "DIGER"):
        raise HTTPException(status_code=400, detail="Geçersiz kategori.")
    conn = get_db()
    ex = conn.execute("SELECT name FROM mappings WHERE code=?", (item.code,)).fetchone()
    name = item.name if item.name else (ex["name"] if ex else item.code)
    conn.execute(
        "INSERT INTO mappings(code,name,category,confirmed) VALUES(?,?,?,1) "
        "ON CONFLICT(code) DO UPDATE SET name=excluded.name, category=excluded.category, confirmed=1",
        (item.code, name, item.category),
    )
    conn.commit()
    conn.close()
    return {"code": item.code, "name": name, "category": item.category, "confirmed": True}


@api_router.delete("/mappings/{code}")
def delete_mapping(code: str):
    conn = get_db()
    conn.execute("DELETE FROM mappings WHERE code=?", (code,))
    conn.commit()
    conn.close()
    return {"ok": True}


@api_router.post("/reset")
def reset_data(req: ResetRequest):
    conn = get_db()
    cur = conn.execute("DELETE FROM reports")
    deleted = cur.rowcount
    mappings_reset = False
    if req.clear_mappings:
        conn.execute("DELETE FROM mappings")
        conn.executemany(
            "INSERT INTO mappings(code,name,category,confirmed) VALUES(?,?,?,1)",
            [(c, n, cat) for c, n, cat in vp.SEED_MAPPINGS],
        )
        mappings_reset = True
    conn.commit()
    conn.close()
    return {"deleted_reports": deleted, "mappings_reset": mappings_reset}


@api_router.post("/reports/{report_id}/export")
def export_report(report_id: str):
    doc = load_report_doc(report_id)
    if not doc:
        raise HTTPException(status_code=404, detail="Rapor bulunamadı.")
    conn = get_db()
    rows = conn.execute("SELECT * FROM mappings").fetchall()
    conn.close()
    doc["mappings"] = [{"code": r["code"], "name": r["name"], "category": r["category"], "confirmed": bool(r["confirmed"])} for r in rows]
    import io
    data = vp.export_excel(doc)
    short = doc.get("period", {}).get("short", "RAPOR")
    filename = f"GIDER_RAPORU_{short}.xlsx".replace(" ", "_")
    headers = {"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"}
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


app.include_router(api_router)


# ----------------------------- Statik arayüz (SPA) -----------------------------
@app.get("/{full_path:path}")
def serve_spa(full_path: str):
    if full_path.startswith("api"):
        raise HTTPException(status_code=404)
    candidate = (STATIC_DIR / full_path).resolve()
    try:
        candidate.relative_to(STATIC_DIR.resolve())
    except ValueError:
        candidate = STATIC_DIR / "index.html"
    if full_path and candidate.is_file():
        return FileResponse(str(candidate))
    index = STATIC_DIR / "index.html"
    if index.is_file():
        return FileResponse(str(index))
    return {"message": "Arayüz dosyaları bulunamadı. Lütfen 'static' klasörünün mevcut olduğundan emin olun."}


def open_browser():
    webbrowser.open(f"http://127.0.0.1:{PORT}")


def main():
    init_db()
    print("=" * 56)
    print("  VEGA GIDER RAPORLAMA - Yerel Masaustu Surumu")
    print(f"  Uygulama baslatiliyor: http://127.0.0.1:{PORT}")
    print(f"  Veriler: {DB_PATH}")
    print("  Kapatmak icin bu pencereyi kapatin.")
    print("=" * 56)
    threading.Timer(1.5, open_browser).start()
    uvicorn.run(app, host="127.0.0.1", port=PORT, log_level="warning")


if __name__ == "__main__":
    main()
