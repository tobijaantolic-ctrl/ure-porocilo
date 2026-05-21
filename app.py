import streamlit as st
import google.generativeai as genai
from pypdf import PdfReader
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import json
import io
import re

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Poročilo opravljenih ur",
    page_icon="📋",
    layout="centered",
)

st.title("📋 Poročilo opravljenih ur")
st.markdown("Naloži **PDF poročilo učitelja** in prejmi Excel datoteko z razčlenjenimi urami.")

# ── Gemini setup ──────────────────────────────────────────────────────────────
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
except Exception:
    st.error("⚠️ Gemini API ključ ni nastavljen. Kontaktiraj administratorja.")
    st.stop()

# ── Helpers: Excel styling ────────────────────────────────────────────────────
def _thin():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)

def _h(cell, text):
    cell.value = text
    cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    cell.fill = PatternFill("solid", start_color="4472C4")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = _thin()

def _c(cell, value, alt=False, fmt=None):
    cell.value = value
    cell.font = Font(name="Arial", size=11)
    cell.fill = PatternFill("solid", start_color="DCE6F1") if alt else PatternFill()
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = _thin()
    if fmt:
        cell.number_format = fmt

def _t(cell, value, fmt=None):
    cell.value = value
    cell.font = Font(bold=True, name="Arial", size=11)
    cell.fill = PatternFill("solid", start_color="FFC000")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = _thin()
    if fmt:
        cell.number_format = fmt

def _widths(ws, widths):
    for col, w in zip("ABCDEFGH", widths):
        ws.column_dimensions[col].width = w

# ── Core: extract PDF text ────────────────────────────────────────────────────
def extract_text(pdf_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    return "\n".join(page.extract_text() or "" for page in reader.pages)

# ── Core: parse with Gemini ───────────────────────────────────────────────────
PROMPT = """
Analiziraj spodnje poročilo učitelja (slovensko šolsko poročilo) in vrni SAMO veljavni JSON brez kakršnegakoli drugega besedila, brez markdown blokov, brez razlag.

JSON mora imeti točno to strukturo:
{{
  "ucitelj": "Ime Priimek",
  "obdobje": "od ... do ...",
  "rap_ps": [
    {{"datum": "D.M.LLLL", "naziv": "RaP PS"}}
  ],
  "dirka_za_branje": [
    {{"datum": "D.M.LLLL", "naziv": "...", "dejavnost": "Sistemizirana"}}
  ],
  "wau": [
    {{"datum": "D.M.LLLL", "opis": "...", "ure": 0.25}}
  ]
}}

Pravila:
- rap_ps: vsi vnosi z nazivom "RaP PS" ali "RAP" (sistemizirana)
- dirka_za_branje: vse OSTALE sistemizirana dejavnosti ki NISO RaP PS/RAP (npr. "Halo Katra", "Urban Šmuc", bralne dirke, prireditve...)
- wau: vsi vnosi ki se začnejo z "wau" IN so nesistemizirana (dežuranje, aktivni odmor itd.)
- Datumi v obliki D.M.LLLL (npr. 3.12.2025)
- Ure za WAU so decimalne (0.25 = 15 minut)

Poročilo:
{text}
"""

def parse_pdf(text: str) -> dict:
    model = genai.GenerativeModel("gemini-2.0-flash")
    response = model.generate_content(PROMPT.format(text=text))
    raw = response.text.strip()
    # Strip markdown fences if present
    raw = re.sub(r"^```[a-z]*\n?", "", raw).rstrip("`").strip()
    return json.loads(raw)

# ── Core: build Excel ─────────────────────────────────────────────────────────
def build_excel(data: dict) -> bytes:
    wb = Workbook()

    rap    = data.get("rap_ps", [])
    dirka  = data.get("dirka_za_branje", [])
    wau    = data.get("wau", [])

    # ── Sheet 1: RAP / RaP PS ─────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "RAP_RaP_PS"
    for col, hdr in enumerate(["Datum", "Naziv", "Ure (upoštevano)"], 1):
        _h(ws1.cell(1, col), hdr)
    for i, row in enumerate(rap):
        r = i + 2
        _c(ws1.cell(r, 1), row.get("datum", ""), i % 2 == 1)
        _c(ws1.cell(r, 2), row.get("naziv", ""), i % 2 == 1)
        _c(ws1.cell(r, 3), 1, i % 2 == 1, fmt="0")
    tr1 = len(rap) + 2
    ws1.merge_cells(f"A{tr1}:B{tr1}")
    _t(ws1.cell(tr1, 1), "SKUPAJ")
    _t(ws1.cell(tr1, 3), f"=SUM(C2:C{tr1-1})", fmt="0")
    _widths(ws1, [14, 18, 20])

    # ── Sheet 2: Dirka za branje ──────────────────────────────────────────────
    ws2 = wb.create_sheet("Dirka_za_branje")
    for col, hdr in enumerate(["Datum", "Naziv", "Dejavnost", "Ure (upoštevano)"], 1):
        _h(ws2.cell(1, col), hdr)
    for i, row in enumerate(dirka):
        r = i + 2
        _c(ws2.cell(r, 1), row.get("datum", ""), i % 2 == 1)
        _c(ws2.cell(r, 2), row.get("naziv", ""), i % 2 == 1)
        _c(ws2.cell(r, 3), row.get("dejavnost", ""), i % 2 == 1)
        _c(ws2.cell(r, 4), 1, i % 2 == 1, fmt="0")
    tr2 = len(dirka) + 2
    ws2.merge_cells(f"A{tr2}:C{tr2}")
    _t(ws2.cell(tr2, 1), "SKUPAJ")
    _t(ws2.cell(tr2, 4), f"=SUM(D2:D{tr2-1})", fmt="0")
    _widths(ws2, [14, 36, 18, 20])

    # ── Sheet 3: WAU ──────────────────────────────────────────────────────────
    ws3 = wb.create_sheet("WAU")
    for col, hdr in enumerate(["Datum", "Opis", "Ure (dejanski čas)"], 1):
        _h(ws3.cell(1, col), hdr)
    for i, row in enumerate(wau):
        r = i + 2
        _c(ws3.cell(r, 1), row.get("datum", ""), i % 2 == 1)
        _c(ws3.cell(r, 2), row.get("opis", ""), i % 2 == 1)
        _c(ws3.cell(r, 3), row.get("ure", 0.25), i % 2 == 1, fmt="0.00")
    tr3 = len(wau) + 2
    ws3.merge_cells(f"A{tr3}:B{tr3}")
    _t(ws3.cell(tr3, 1), "SKUPAJ")
    _t(ws3.cell(tr3, 3), f"=SUM(C2:C{tr3-1})", fmt="0.00")
    _widths(ws3, [14, 28, 22])

    # ── Sheet 4: Poročilo o urah ───────────────────────────────────────────────
    ws4 = wb.create_sheet("Poročilo o urah")
    for col, hdr in enumerate(["Tabela", "Ure"], 1):
        _h(ws4.cell(1, col), hdr)
    summary = [
        ("RAP / RaP PS",    f"=RAP_RaP_PS!C{tr1}"),
        ("Dirka za branje", f"=Dirka_za_branje!D{tr2}"),
        ("WAU",             f"=WAU!C{tr3}"),
    ]
    for i, (label, formula) in enumerate(summary):
        r = i + 2
        _c(ws4.cell(r, 1), label, i % 2 == 1)
        _c(ws4.cell(r, 2), formula, i % 2 == 1, fmt="0.00")
    _t(ws4.cell(5, 1), "SKUPAJ")
    _t(ws4.cell(5, 2), "=SUM(B2:B4)", fmt="0.00")
    _widths(ws4, [24, 14])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ── UI ────────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader("📄 Naloži PDF poročilo", type=["pdf"])

if uploaded:
    st.info(f"Datoteka: **{uploaded.name}**")

    if st.button("⚙️ Obdelaj in ustvari Excel", type="primary"):
        with st.spinner("Berem PDF..."):
            pdf_bytes = uploaded.read()
            text = extract_text(pdf_bytes)

        with st.spinner("Analiziram z AI (Gemini)..."):
            try:
                data = parse_pdf(text)
            except Exception as e:
                st.error(f"Napaka pri analizi PDF-ja: {e}")
                st.text_area("Surovo besedilo PDF-ja (za debug):", text[:3000])
                st.stop()

        with st.spinner("Ustvarjam Excel..."):
            excel_bytes = build_excel(data)

        # ── Summary ───────────────────────────────────────────────────────────
        ucitelj = data.get("ucitelj", "Neznano")
        obdobje = data.get("obdobje", "")
        n_rap   = len(data.get("rap_ps", []))
        n_dirka = len(data.get("dirka_za_branje", []))
        n_wau   = len(data.get("wau", []))
        wau_h   = sum(r.get("ure", 0) for r in data.get("wau", []))

        st.success("✅ Excel je pripravljen!")
        col1, col2 = st.columns(2)
        col1.markdown(f"**Učitelj:** {ucitelj}")
        col2.markdown(f"**Obdobje:** {obdobje}")

        col3, col4, col5 = st.columns(3)
        col3.metric("RAP / RaP PS", f"{n_rap} ur")
        col4.metric("Dirka za branje", f"{n_dirka} ur")
        col5.metric("WAU", f"{wau_h:.2f} ur")

        # ── Download ──────────────────────────────────────────────────────────
        safe_name = ucitelj.replace(" ", "_").replace("/", "-")
        st.download_button(
            label="⬇️ Prenesi Excel",
            data=excel_bytes,
            file_name=f"porocilo_{safe_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.markdown("---")
st.caption("Šmarje-Sap · Poročilo opravljenih ur · Powered by Gemini AI")
