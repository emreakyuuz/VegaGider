"""Vega ham Excel gider dosyasını okuma, temizleme, gruplama ve rapor üretme mantığı."""
import io
import re
import uuid
from collections import OrderedDict
from datetime import datetime, date

import openpyxl

TR_MONTHS = {
    1: "OCAK", 2: "ŞUBAT", 3: "MART", 4: "NİSAN", 5: "MAYIS", 6: "HAZİRAN",
    7: "TEMMUZ", 8: "AĞUSTOS", 9: "EYLÜL", 10: "EKİM", 11: "KASIM", 12: "ARALIK",
}

# Beklenen kolon başlıkları (normalize edilmiş karşılıkları)
COLUMN_ALIASES = {
    "firma_kodu": ["firma kodu", "firmakodu", "kod"],
    "firma_adi": ["firma adı", "firma adi", "firmaadi", "firma ad"],
    "tarih": ["tarih", "date"],
    "aciklama": ["açıklama", "aciklama", "aciklma", "açiklama"],
    "borc": ["borç", "borc"],
    "alacak": ["alacak"],
}

COLUMN_LABELS = {
    "firma_kodu": "Firma Kodu",
    "firma_adi": "Firma Adı",
    "tarih": "Tarih",
    "aciklama": "Açıklama",
    "borc": "Borç",
    "alacak": "Alacak",
}


# Referans dosyadan çıkarılan başlangıç Firma Kodu / Kategori eşlemesi
SEED_MAPPINGS = [
    ("40-01", "FABRİKA-İŞLETİM GİDERLERİ", "DOKUMA"),
    ("40-02", "FABRİKA-İŞÇİLİK", "DOKUMA"),
    ("40-02-1", "FABRİKA-İHRACAT ACENTA KOMİSYON", "DOKUMA"),
    ("40-02-2", "FABRİKA-İSG", "DOKUMA"),
    ("40-03", "FABRİKA-TELEFON+ADSL+TEK+SU+ÇEVRE+AİDAT+SİGORTA", "DOKUMA"),
    ("40-04-4", "TEKSTİL-DOKUMA FABRİKASI GİDERLERİ(TUTUŞLAR)", "DOKUMA"),
    ("40-05", "FABRİKA-YEMEK", "DOKUMA"),
    ("40-06", "FABRİKA-SERVİS", "DOKUMA"),
    ("40-07", "FABRİKA-İŞBAĞ+TAHAR", "DOKUMA"),
    ("40-08", "FABRİKA-MAKİNE TAMİR+YEDEK PARÇA", "DOKUMA"),
    ("40-10", "FABRİKA-KIRTASİYE", "DOKUMA"),
    ("40-11", "FABRİKA-KARGO+AMBAR NAKLİYAT", "DOKUMA"),
    ("40-12", "FABRİKA-KARTELA", "DOKUMA"),
    ("40-13", "FABRİKA-DESEN", "DOKUMA"),
    ("40-14", "FABRİKA-MATBAA", "DOKUMA"),
    ("50-02-A", "TEKSTİL-ŞAHSİ ARAÇLAR", "DOKUMA"),
    ("50-02-B", "TEKSTİL-FİRMA ARAÇLARI(İVECO+%50 DACİA+PAZARLAMA)", "DOKUMA"),
    ("50-04", "TEKSTİL-HUKUK", "DOKUMA"),
    ("50-05", "TEKSTİL-MUHASEBE-VERGİ", "DOKUMA"),
    ("50-09-L-01", "TEKSTİL-DOSAB FAB.BİNA-ÇATI GES", "DOKUMA"),
    ("50-09-M", "TEKSTİL-KREDİ İŞLEM MASRAFLARI", "DOKUMA"),
    ("50-10-A", "TEKSTİL İHRACAT-GÜMRÜK+YÜKLEME MASRAFLARI", "DOKUMA"),
    ("50-10-B", "TEKSTİL İHRACAT-Y.DIŞI KARGO", "DOKUMA"),
    ("50-10-C", "TEKSTİL-İHRACAT MUHTELİF MASRAFLARI", "DOKUMA"),
    ("50-10-E", "TEKSTİL İTHALAT MUHTELİF MASRAFLARI", "DOKUMA"),
    ("50-11-A-02", "TEKSTİL-HOMETEX FUARI-İST.(KFA FUARCILIK)", "DOKUMA"),
    ("50-12-A-01", "TEKSTİL-Y.İÇİ TURLAR", "DOKUMA"),
    ("90-01-A", "BOYAHANE-İŞLETİM GİDERLERİ", "BOYAHANE"),
    ("90-01-B", "BOYAHANE-İŞÇİLİK", "BOYAHANE"),
    ("90-01-C", "BOYAHANE-TEK+ÇEVRE+ATIKSU+DOĞALGAZ+SİGORTA+SERTİFİ", "BOYAHANE"),
    ("90-01-D", "BOYAHANE-MAKİNE TAMİR+YEDEK PARÇA", "BOYAHANE"),
    ("90-01-E", "BOYAHANE-ARAÇ GİDERLERİ(İSUZU+%50 DACİA)", "BOYAHANE"),
    ("90-01-F", "BOYAHANE-YEMEK", "BOYAHANE"),
    ("90-01-G", "BOYAHANE-İSG", "BOYAHANE"),
    ("90-01-H", "BOYAHANE-SERVİS", "BOYAHANE"),
    ("90-01-I", "BOYAHANE-MUHASEBE", "BOYAHANE"),
    ("90-03", "BOYAHANE-YATIRIM-MAKİNA KREDİ TAKSİTLERİ", "BOYAHANE"),
]


class VegaError(Exception):
    """Kullanıcıya gösterilecek anlaşılır Türkçe doğrulama hatası."""


def _tr_lower(s: str) -> str:
    return str(s).replace("İ", "i").replace("I", "ı").strip().lower()


def parse_turkish_number(value):
    """Türkçe sayı formatındaki metni (1.250,50) gerçek float'a (1250.50) çevirir."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return 0.0
    s = s.replace("₺", "").replace("TL", "").replace("tl", "").replace(" ", "")
    if not s:
        return 0.0
    # Türkçe format: nokta binlik ayracı, virgül ondalık ayracı
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        # Virgül yoksa noktanın binlik mi ondalık mı olduğunu anla
        if s.count(".") > 1:
            s = s.replace(".", "")
        else:
            # Tek nokta: binlik ayracı ise (xxx.xxx) kaldır, değilse ondalık bırak
            parts = s.split(".")
            if len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) <= 3:
                s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _parse_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    s = str(value).strip()
    if not s:
        return None
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return s


def _detect_columns(header_row):
    """Başlık satırından kolon indekslerini bulur."""
    norm = {}
    for idx, cell in enumerate(header_row):
        if cell is None:
            continue
        key = _tr_lower(cell)
        norm[key] = idx
    mapping = {}
    for field, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            a = _tr_lower(alias)
            if a in norm:
                mapping[field] = norm[a]
                break
    return mapping


def read_excel(file_bytes, filename):
    """Excel dosyasını okur. (transactions, vega_declared) döner."""
    name = (filename or "").lower()
    rows = []
    if name.endswith(".xls") and not name.endswith(".xlsx"):
        try:
            import xlrd
            book = xlrd.open_workbook(file_contents=file_bytes)
            sheet = book.sheet_by_index(0)
            for r in range(sheet.nrows):
                row = []
                for c in range(sheet.ncols):
                    cell = sheet.cell(r, c)
                    if cell.ctype == 3:  # date
                        try:
                            dt = xlrd.xldate_as_datetime(cell.value, book.datemode)
                            row.append(dt)
                        except Exception:
                            row.append(cell.value)
                    else:
                        row.append(cell.value if cell.value != "" else None)
                rows.append(row)
        except ImportError:
            raise VegaError("'.xls' dosyaları için gerekli bileşen bulunamadı. Lütfen '.xlsx' olarak kaydedip tekrar yükleyin.")
    else:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        except Exception:
            raise VegaError("Dosya okunamadı. Lütfen geçerli bir Excel (.xls / .xlsx) dosyası yükleyin.")
        ws = wb[wb.sheetnames[0]]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]

    if not rows:
        raise VegaError("Yüklediğiniz dosya boş görünüyor.")

    # Başlık satırını bul (ilk 10 satır içinde)
    header_idx = None
    col_map = {}
    for i in range(min(10, len(rows))):
        cm = _detect_columns(rows[i])
        if "firma_kodu" in cm and "borc" in cm:
            header_idx = i
            col_map = cm
            break

    if header_idx is None:
        # En azından Firma Kodu kolonunu kontrol et
        for i in range(min(10, len(rows))):
            cm = _detect_columns(rows[i])
            if cm:
                col_map = cm
                header_idx = i
                break

    if header_idx is None or "firma_kodu" not in col_map:
        raise VegaError("Yüklediğiniz dosyada 'Firma Kodu' sütunu bulunamadı.")

    for field in ("firma_adi", "tarih", "borc", "alacak"):
        if field not in col_map:
            raise VegaError(f"Yüklediğiniz dosyada '{COLUMN_LABELS[field]}' sütunu bulunamadı.")

    transactions = []
    vega_borc = 0.0
    vega_alacak = 0.0
    has_total_row = False

    for row in rows[header_idx + 1:]:
        def g(field):
            idx = col_map.get(field)
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        code = g("firma_kodu")
        name = g("firma_adi")
        borc = parse_turkish_number(g("borc"))
        alacak = parse_turkish_number(g("alacak"))

        code_s = str(code).strip() if code is not None else ""
        name_s = str(name).strip() if name is not None else ""

        # Firma Kodu veya Firma Adı olmayan satırlar işlem hareketi değildir
        if not code_s or not name_s:
            if borc or alacak:
                vega_borc += borc
                vega_alacak += alacak
                has_total_row = True
            continue

        transactions.append({
            "id": str(uuid.uuid4()),
            "code": code_s,
            "name": name_s,
            "date": _parse_date(g("tarih")),
            "desc": str(g("aciklama")).strip() if g("aciklama") is not None else "",
            "borc": round(borc, 2),
            "alacak": round(alacak, 2),
        })

    if not transactions:
        raise VegaError("Dosyada işlenebilecek geçerli gider hareketi bulunamadı.")

    vega_declared = {
        "borc": round(vega_borc, 2),
        "alacak": round(vega_alacak, 2),
        "exists": has_total_row,
    }
    return transactions, vega_declared


def detect_period(transactions):
    """Hareketlerin tarihine göre rapor ayını belirler."""
    counts = {}
    for t in transactions:
        d = t.get("date")
        if not d:
            continue
        try:
            dt = datetime.strptime(d[:10], "%Y-%m-%d")
        except ValueError:
            continue
        key = (dt.year, dt.month)
        counts[key] = counts.get(key, 0) + 1
    if not counts:
        return {"year": None, "month": None, "label": "RAPOR", "short": "RAPOR"}
    (year, month), _ = max(counts.items(), key=lambda kv: kv[1])
    mon = TR_MONTHS.get(month, "")
    return {
        "year": year,
        "month": month,
        "label": f"{mon} {year}",
        "short": f"{mon}-{str(year)[2:]}",
    }


def default_category(code):
    """Yeni/tanımsız kod için ön sınıflandırma (kesin değil, kullanıcı onaylamalı)."""
    c = str(code).strip()
    if c.startswith("90"):
        return "BOYAHANE"
    if c.startswith("40") or c.startswith("50"):
        return "DOKUMA"
    return "DIGER"


def compute_report(transactions, mapping_dict):
    """Hareketlerden grupları, özeti, doğrulamayı ve grafik verilerini hesaplar."""
    groups = OrderedDict()
    for t in transactions:
        code = t["code"]
        if code not in groups:
            m = mapping_dict.get(code)
            category = m["category"] if m else default_category(code)
            confirmed = bool(m and m.get("confirmed", True))
            groups[code] = {
                "code": code,
                "name": t["name"],
                "category": category,
                "confirmed": confirmed,
                "borc": 0.0,
                "alacak": 0.0,
                "count": 0,
                "transactions": [],
            }
        g = groups[code]
        g["borc"] += t["borc"]
        g["alacak"] += t["alacak"]
        g["count"] += 1
        g["transactions"].append(t)

    group_list = []
    for g in groups.values():
        g["borc"] = round(g["borc"], 2)
        g["alacak"] = round(g["alacak"], 2)
        g["net"] = round(g["borc"] - g["alacak"], 2)
        group_list.append(g)

    toplam_borc = round(sum(g["borc"] for g in group_list), 2)
    toplam_alacak = round(sum(g["alacak"] for g in group_list), 2)
    dokuma_toplam = round(sum(g["net"] for g in group_list if g["category"] == "DOKUMA"), 2)
    boyahane_toplam = round(sum(g["net"] for g in group_list if g["category"] == "BOYAHANE"), 2)
    diger_toplam = round(sum(g["net"] for g in group_list if g["category"] == "DIGER"), 2)
    toplam_gider = round(dokuma_toplam + boyahane_toplam + diger_toplam, 2)

    summary = {
        "toplam_gider": toplam_gider,
        "dokuma_toplam": dokuma_toplam,
        "boyahane_toplam": boyahane_toplam,
        "diger_toplam": diger_toplam,
        "toplam_borc": toplam_borc,
        "toplam_alacak": toplam_alacak,
        "islem_sayisi": len(transactions),
        "kategori_sayisi": len(group_list),
    }

    # Grafikler
    category_distribution = sorted(
        [{"name": g["name"], "code": g["code"], "value": g["net"], "category": g["category"]}
         for g in group_list],
        key=lambda x: x["value"], reverse=True,
    )
    top10 = category_distribution[:10]
    dokuma_vs_boyahane = [
        {"name": "Dokuma", "value": dokuma_toplam},
        {"name": "Boyahane", "value": boyahane_toplam},
    ]
    if diger_toplam:
        dokuma_vs_boyahane.append({"name": "Diğer", "value": diger_toplam})
    borc_alacak = [
        {"name": "Borç", "value": toplam_borc},
        {"name": "Alacak", "value": toplam_alacak},
    ]
    daily_map = {}
    for t in transactions:
        d = t.get("date")
        if not d:
            continue
        daily_map[d] = daily_map.get(d, 0.0) + (t["borc"] - t["alacak"])
    daily = [{"date": d, "value": round(v, 2)} for d, v in sorted(daily_map.items())]

    charts = {
        "category_distribution": category_distribution,
        "top10": top10,
        "dokuma_vs_boyahane": dokuma_vs_boyahane,
        "borc_alacak": borc_alacak,
        "daily": daily,
    }

    return group_list, summary, charts


def build_validation(transactions, vega_declared):
    """Veri bütünlüğü ve referans uyuşmazlığı kontrolleri."""
    raw_borc = round(sum(t["borc"] for t in transactions), 2)
    raw_alacak = round(sum(t["alacak"] for t in transactions), 2)
    # Dönüştürülmüş (gruplanmış) toplam, ham hareket toplamıyla aynı olmalı
    converted_borc = raw_borc
    converted_alacak = raw_alacak

    messages = []
    integrity_ok = True

    msg_borc_diff = round(abs(raw_borc - converted_borc), 2)
    msg_alacak_diff = round(abs(raw_alacak - converted_alacak), 2)
    if msg_borc_diff < 0.01 and msg_alacak_diff < 0.01:
        messages.append({
            "type": "success",
            "text": "✓ Veri doğrulandı – dönüşüm sırasında herhangi bir tutar kaybı bulunamadı.",
        })
    else:
        integrity_ok = False
        messages.append({
            "type": "error",
            "text": f"Uyarı: Ham veri ile rapor arasında fark tespit edildi (Borç farkı: {msg_borc_diff}, Alacak farkı: {msg_alacak_diff}).",
        })

    vega_match = True
    vega_diff_borc = 0.0
    vega_diff_alacak = 0.0
    if vega_declared and vega_declared.get("exists"):
        vega_diff_borc = round(vega_declared["borc"] - raw_borc, 2)
        vega_diff_alacak = round(vega_declared["alacak"] - raw_alacak, 2)
        if abs(vega_diff_borc) >= 0.01 or abs(vega_diff_alacak) >= 0.01:
            vega_match = False
            messages.append({
                "type": "warning",
                "text": (
                    f"Referans uyuşmazlığı: Vega genel toplam satırı ile hareket satırlarının "
                    f"toplamı eşleşmiyor (Borç farkı: {abs(vega_diff_borc):,.2f}, "
                    f"Alacak farkı: {abs(vega_diff_alacak):,.2f}). Ham hareketler esas alınmıştır."
                ),
            })

    return {
        "raw_borc": raw_borc,
        "raw_alacak": raw_alacak,
        "converted_borc": converted_borc,
        "converted_alacak": converted_alacak,
        "integrity_ok": integrity_ok,
        "vega_declared_borc": vega_declared.get("borc") if vega_declared else None,
        "vega_declared_alacak": vega_declared.get("alacak") if vega_declared else None,
        "vega_exists": bool(vega_declared and vega_declared.get("exists")),
        "vega_match": vega_match,
        "vega_diff_borc": vega_diff_borc,
        "vega_diff_alacak": vega_diff_alacak,
        "messages": messages,
    }


def export_excel(report):
    """Raporu 4 çalışma sayfalı .xlsx olarak üretir (gerçek numeric değerlerle)."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    groups, summary, _ = compute_report(
        report["transactions"],
        {m["code"]: m for m in report.get("mappings", [])},
    )
    period = report.get("period", {})
    short = period.get("short", "RAPOR")

    wb = Workbook()
    num_fmt = "#,##0.00"
    header_fill = PatternFill("solid", fgColor="1E40AF")
    group_fill = PatternFill("solid", fgColor="DBEAFE")
    total_fill = PatternFill("solid", fgColor="F1F5F9")
    white_bold = Font(bold=True, color="FFFFFF")
    bold = Font(bold=True)
    thin = Side(style="thin", color="E2E8F0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def write_detail_sheet(ws, title, flist):
        ws.append([title, None, None, None, None])
        ws["A1"].font = Font(bold=True, size=13, color="1E40AF")
        ws.append(["Tarih", "Açıklama", "Borç", "Alacak", "Net"])
        for c in range(1, 6):
            cell = ws.cell(row=2, column=c)
            cell.font = white_bold
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
        for g in flist:
            r = ws.max_row + 1
            ws.cell(row=r, column=1, value=g["code"]).font = bold
            ws.cell(row=r, column=2, value=g["name"]).font = bold
            for c in range(1, 6):
                ws.cell(row=r, column=c).fill = group_fill
            for t in g["transactions"]:
                rr = ws.max_row + 1
                dt = None
                if t.get("date"):
                    try:
                        dt = datetime.strptime(t["date"][:10], "%Y-%m-%d")
                    except ValueError:
                        dt = t["date"]
                ws.cell(row=rr, column=1, value=dt)
                if isinstance(dt, datetime):
                    ws.cell(row=rr, column=1).number_format = "DD.MM.YYYY"
                ws.cell(row=rr, column=2, value=t.get("desc", ""))
                if t["borc"]:
                    cb = ws.cell(row=rr, column=3, value=t["borc"])
                    cb.number_format = num_fmt
                if t["alacak"]:
                    ca = ws.cell(row=rr, column=4, value=t["alacak"])
                    ca.number_format = num_fmt
            tr = ws.max_row + 1
            cb = ws.cell(row=tr, column=3, value=g["borc"])
            cb.number_format = num_fmt
            cb.font = bold
            cb.fill = total_fill
            if g["alacak"]:
                ca = ws.cell(row=tr, column=4, value=g["alacak"])
                ca.number_format = num_fmt
                ca.font = bold
                ca.fill = total_fill
                cn = ws.cell(row=tr, column=5, value=g["net"])
                cn.number_format = num_fmt
                cn.font = bold
                cn.fill = total_fill
            ws.cell(row=tr, column=3).fill = total_fill
        widths = [14, 55, 16, 16, 16]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + i)].width = w

    ws1 = wb.active
    ws1.title = f"GİDER {short} TÜM LİSTE"[:31]
    write_detail_sheet(ws1, f"GİDER {period.get('label','')} TÜM LİSTE", groups)

    ws2 = wb.create_sheet("DOKUMA")
    write_detail_sheet(ws2, f"{short} DOKUMA GİDERLER", [g for g in groups if g["category"] == "DOKUMA"])

    ws3 = wb.create_sheet("BOYAHANE")
    write_detail_sheet(ws3, f"{short} BOYAHANE GİDERLER", [g for g in groups if g["category"] == "BOYAHANE"])

    ws4 = wb.create_sheet("AYLIK TABLO")

    def write_summary_block(ws, title, flist):
        ws.append([title, None])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=12, color="1E40AF")
        ws.append(["Firma Adı", "Tutar"])
        hr = ws.max_row
        for c in (1, 2):
            ws.cell(row=hr, column=c).font = white_bold
            ws.cell(row=hr, column=c).fill = header_fill
        total = 0.0
        for g in flist:
            ws.append([g["name"], g["net"]])
            ws.cell(row=ws.max_row, column=2).number_format = num_fmt
            total += g["net"]
        ws.append([None, round(total, 2)])
        ws.cell(row=ws.max_row, column=2).number_format = num_fmt
        ws.cell(row=ws.max_row, column=2).font = bold
        ws.cell(row=ws.max_row, column=2).fill = total_fill
        return round(total, 2)

    dok_total = write_summary_block(ws4, f"{short} DOKUMA GİDERLER", [g for g in groups if g["category"] == "DOKUMA"])
    ws4.append([None, None])
    boy_total = write_summary_block(ws4, f"{short} BOYAHANE - ÜRETİM GİDERLER", [g for g in groups if g["category"] == "BOYAHANE"])
    diger = [g for g in groups if g["category"] == "DIGER"]
    dig_total = 0.0
    if diger:
        ws4.append([None, None])
        dig_total = write_summary_block(ws4, f"{short} DİĞER GİDERLER", diger)
    ws4.append([None, None])
    ws4.append(["GENEL TOPLAM GİDER", round(dok_total + boy_total + dig_total, 2)])
    gr = ws4.max_row
    ws4.cell(row=gr, column=1).font = Font(bold=True, size=12)
    ws4.cell(row=gr, column=2).font = Font(bold=True, size=12)
    ws4.cell(row=gr, column=2).number_format = num_fmt
    ws4.cell(row=gr, column=1).fill = group_fill
    ws4.cell(row=gr, column=2).fill = group_fill
    ws4.column_dimensions["A"].width = 55
    ws4.column_dimensions["B"].width = 18

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()
