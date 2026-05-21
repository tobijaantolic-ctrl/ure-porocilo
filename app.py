import streamlit as st
from pypdf import PdfReader
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re
import io

st.set_page_config(page_title="Poročilo opravljenih ur", page_icon="📋", layout="centered")
st.title("📋 Poročilo opravljenih ur")
st.markdown("Naloži **PDF poročilo učitelja** in prejmi Excel datoteko z razčlenjenimi urami.")

# ── Excel styling helpers ─────────────────────────────────────────────────────
def _border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)

def _h(cell, text):
    cell.value = text
    cell.font  = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    cell.fill  = PatternFill("solid", start_color="4472C4")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = _border()

def _c(cell, value, alt=False, fmt=None):
    cell.value = value
    cell.font  = Font(name="Arial", size=11)
    cell.fill  = PatternFill("solid", start_color="DCE6F1") if alt else PatternFill()
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = _border()
    if fmt: cell.number_format = fmt

def _t(cell, value, fmt=None):
    cell.value = value
    cell.font  = Font(bold=True, name="Arial", size=11)
    cell.fill  = PatternFill("solid", start_color="FFC000")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = _border()
    if fmt: cell.number_format = fmt

def _widths(ws, widths):
    for col, w in zip("ABCDEFGH", widths):
        ws.column_dimensions[col].width = w

# ── PDF parsing (no API needed) ───────────────────────────────────────────────
def extract_text(pdf_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    return "\n".join(page.extract_text() or "" for page in reader.pages)

def parse_pdf(text: str) -> dict:
    # Extract teacher name
    ucitelj = "Neznano"
    m = re.search(r'Poročilo učitelja\s*\n([^\n]+)', text)
    if m:
        ucitelj = m.group(1).strip()

    # Extract period
    obdobje = ""
    m = re.search(r'Obdobje:\s*(od\s+[\d\.\s]+do\s+[\d\.\s]+)', text)
    if m:
        obdobje = m.group(1).strip()

    # Match every activity row:
    # Format: D. M. YYYY  naziv  group.  0h Xm (decimal)  Sistemizirana|Nesistemizirana
    pattern = re.compile(
        r'(\d{1,2}\.\s*\d{1,2}\.\s*\d{4})'   # date
        r'\s+(.+?)'                             # naziv (non-greedy)
        r'\s+\d+\.'                             # group number
        r'\s+\d+h\s+\d+m'                      # time
        r'\s+\((\d+[.,]\d+)\)'                 # (decimal hours)
        r'\s+(Sistemizirana|Nesistemizirana)'   # type
    )

    rap_ps, dirka, wau_list = [], [], []

    for m in pattern.finditer(text):
        raw_datum, naziv, ure_str, dejavnost = m.groups()
        datum = re.sub(r'\s+', '', raw_datum)   # "2. 12. 2025" → "2.12.2025"
        ure   = float(ure_str.replace(',', '.'))
        naziv = naziv.strip()

        if re.search(r'\brap\s*ps\b|\brap\b', naziv, re.IGNORECASE):
            rap_ps.append({"datum": datum, "naziv": naziv})
        elif naziv.lower().startswith("wau"):
            wau_list.append({"datum": datum, "opis": naziv[4:].strip(), "ure": ure})
        elif dejavnost == "Sistemizirana":
            dirka.append({"datum": datum, "naziv": naziv, "dejavnost": dejavnost})
        else:
            wau_list.append({"datum": datum, "opis": naziv, "ure": ure})

    return {
        "ucitelj": ucitelj,
        "obdobje": obdobje,
        "rap_ps": rap_ps,
        "dirka_za_branje": dirka,
        "wau": wau_list,
    }

# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(data: dict) -> bytes:
    wb  = Workbook()
    rap   = data["rap_ps"]
    dirka = data["dirka_za_branje"]
    wau   = data["wau"]

    # Sheet 1 — RAP / RaP PS
    ws1 = wb.active; ws1.title = "RAP_RaP_PS"
    for c, h in enumerate(["Datum", "Naziv", "Ure (upoštevano)"], 1): _h(ws1.cell(1, c), h)
    for i, r in enumerate(rap):
        _c(ws1.cell(i+2, 1), r["datum"],  i%2); _c(ws1.cell(i+2, 2), r["naziv"], i%2)
        _c(ws1.cell(i+2, 3), 1, i%2, "0")
    tr1 = len(rap)+2
    ws1.merge_cells(f"A{tr1}:B{tr1}"); _t(ws1.cell(tr1,1), "SKUPAJ")
    _t(ws1.cell(tr1, 3), f"=SUM(C2:C{tr1-1})", "0")
    _widths(ws1, [14, 22, 20])

    # Sheet 2 — Dirka za branje
    ws2 = wb.create_sheet("Dirka_za_branje")
    for c, h in enumerate(["Datum", "Naziv", "Dejavnost", "Ure (upoštevano)"], 1): _h(ws2.cell(1,c), h)
    for i, r in enumerate(dirka):
        _c(ws2.cell(i+2,1), r["datum"],      i%2); _c(ws2.cell(i+2,2), r["naziv"],      i%2)
        _c(ws2.cell(i+2,3), r["dejavnost"],  i%2); _c(ws2.cell(i+2,4), 1, i%2, "0")
    tr2 = len(dirka)+2
    ws2.merge_cells(f"A{tr2}:C{tr2}"); _t(ws2.cell(tr2,1), "SKUPAJ")
    _t(ws2.cell(tr2,4), f"=SUM(D2:D{tr2-1})", "0")
    _widths(ws2, [14, 36, 18, 20])

    # Sheet 3 — WAU
    ws3 = wb.create_sheet("WAU")
    for c, h in enumerate(["Datum", "Opis", "Ure (dejanski čas)"], 1): _h(ws3.cell(1,c), h)
    for i, r in enumerate(wau):
        _c(ws3.cell(i+2,1), r["datum"], i%2); _c(ws3.cell(i+2,2), r["opis"],  i%2)
        _c(ws3.cell(i+2,3), r["ure"],   i%2, "0.00")
    tr3 = len(wau)+2
    ws3.merge_cells(f"A{tr3}:B{tr3}"); _t(ws3.cell(tr3,1), "SKUPAJ")
    _t(ws3.cell(tr3,3), f"=SUM(C2:C{tr3-1})", "0.00")
    _widths(ws3, [14, 30, 22])

    # Sheet 4 — Poročilo o urah
    ws4 = wb.create_sheet("Poročilo o urah")
    for c, h in enumerate(["Tabela", "Ure"], 1): _h(ws4.cell(1,c), h)
    for i, (lbl, fml) in enumerate([
        ("RAP / RaP PS",    f"=RAP_RaP_PS!C{tr1}"),
        ("Dirka za branje", f"=Dirka_za_branje!D{tr2}"),
        ("WAU",             f"=WAU!C{tr3}"),
    ]):
        _c(ws4.cell(i+2,1), lbl, i%2); _c(ws4.cell(i+2,2), fml, i%2, "0.00")
    _t(ws4.cell(5,1), "SKUPAJ"); _t(ws4.cell(5,2), "=SUM(B2:B4)", "0.00")
    _widths(ws4, [24, 14])

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

# ── UI ────────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader("📄 Naloži PDF poročilo", type=["pdf"])

if uploaded:
    st.info(f"Datoteka: **{uploaded.name}**")
    if st.button("⚙️ Obdelaj in ustvari Excel", type="primary"):

        with st.spinner("Berem in analiziram PDF..."):
            text = extract_text(uploaded.read())
            data = parse_pdf(text)

        rap_n   = len(data["rap_ps"])
        dirka_n = len(data["dirka_za_branje"])
        wau_n   = len(data["wau"])
        wau_h   = sum(r["ure"] for r in data["wau"])

        if rap_n + dirka_n + wau_n == 0:
            st.error("Ni bilo mogoče prepoznati aktivnosti v PDF-ju. Pokaži surovo besedilo:")
            st.text_area("Surovo besedilo:", text[:3000], height=200)
        else:
            excel_bytes = build_excel(data)
            st.success("✅ Excel je pripravljen!")

            c1, c2 = st.columns(2)
            c1.markdown(f"**Učitelj:** {data['ucitelj']}")
            c2.markdown(f"**Obdobje:** {data['obdobje']}")

            c3, c4, c5 = st.columns(3)
            c3.metric("RAP / RaP PS", f"{rap_n} ur")
            c4.metric("Dirka za branje", f"{dirka_n} ur")
            c5.metric("WAU", f"{wau_h:.2f} ur")

            safe = data["ucitelj"].replace(" ", "_")
            st.download_button(
                "⬇️ Prenesi Excel", excel_bytes,
                file_name=f"porocilo_{safe}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

st.markdown("---")
st.caption("Šmarje-Sap · Poročilo opravljenih ur · Brez AI API")
