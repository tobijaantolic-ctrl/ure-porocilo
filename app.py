        line = lines[i]
        if not DATE_LINE_RE.match(line):
            i += 1
            continue

        row = line
        i += 1
        while i < len(lines) and not TIME_ACTIVITY_RE.search(row):
            next_line = lines[i]
            if DATE_LINE_RE.match(next_line):
                break
            row = f"{row} {next_line}"
            i += 1

        rows.append(row)

    return "\n".join(rows)

def parse_pdf(text: str) -> dict:
    ucitelj = "Neznano"
    m = re.search(r'Poročilo učitelja\s*\n([^\n]+)', text)
    if m: ucitelj = m.group(1).strip()

    obdobje = ""
    m = re.search(r'Obdobje:\s*(od\s+[\d\.\s]+do\s+[\d\.\s]+)', text)
    if m: obdobje = m.group(1).strip()

    rap_ps, dirka, wau_list = [], [], []

    for m in PATTERN.finditer(activity_lines(text)):
        raw_datum, raw_naziv, ure_str, dejavnost = m.groups()
        datum     = re.sub(r'\s+', '', raw_datum)
        ure       = float(ure_str.replace(',', '.'))
        naziv_low = raw_naziv.strip().lower()
        is_wau    = naziv_low.startswith("wau")
        is_rap    = bool(re.search(r'\brap\b', naziv_low))

        if is_wau:
            # WAU = SAMO dejavnosti z "wau" predpono → dejanski čas
            wau_list.append({
                "datum": datum,
                "opis":  clean_naziv(raw_naziv.strip()),
                "ure":   ure,
            })
        elif is_rap:
            # RAP / RaP PS / RaP RS → pedagoška ura, zaokroženo navzgor
            rap_ps.append({
                "datum":  datum,
                "naziv":  rap_naziv(raw_naziv.strip()),
                "ure_st": ped_ure(ure),
            })
        else:
            # VSE ostalo (sistemizirana ALI nesistemizirana, brez "wau") → pedagoška ura
            dirka.append({
                "datum":     datum,
                "naziv":     clean_naziv(raw_naziv.strip(), strip_od_do=True),
                "dejavnost": dejavnost,
                "ure_st":    ped_ure(ure),
            })

    return {
        "ucitelj": ucitelj,
        "obdobje": obdobje,
        "rap_ps":  rap_ps,
        "dirka_za_branje": dirka,
        "wau":     wau_list,
    }

# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(data: dict) -> bytes:
    wb    = Workbook()
    rap   = data["rap_ps"]
    dirka = data["dirka_za_branje"]
    wau   = data["wau"]

    # Sheet 1 — RAP / RaP PS / RaP RS
    ws1 = wb.active; ws1.title = "RAP_RaP_PS"
    for c, h in enumerate(["Datum", "Naziv", "Ure (upoštevano)"], 1): _h(ws1.cell(1,c), h)
    for i, r in enumerate(rap):
        _c(ws1.cell(i+2,1), r["datum"],  i%2)
        _c(ws1.cell(i+2,2), r["naziv"],  i%2)
        _c(ws1.cell(i+2,3), r["ure_st"], i%2, "0")
    tr1 = len(rap)+2
    ws1.merge_cells(f"A{tr1}:B{tr1}"); _t(ws1.cell(tr1,1), "SKUPAJ")
    _t(ws1.cell(tr1,3), f"=SUM(C2:C{tr1-1})", "0")
    _widths(ws1, [14, 22, 20])

    # Sheet 2 — Dirka za branje / vse ostale pedagoške ure
    ws2 = wb.create_sheet("Dirka_za_branje")
    for c, h in enumerate(["Datum", "Naziv", "Dejavnost", "Ure (upoštevano)"], 1): _h(ws2.cell(1,c), h)
    for i, r in enumerate(dirka):
        _c(ws2.cell(i+2,1), r["datum"],     i%2)
        _c(ws2.cell(i+2,2), r["naziv"],     i%2)
        _c(ws2.cell(i+2,3), r["dejavnost"], i%2)
        _c(ws2.cell(i+2,4), r["ure_st"],    i%2, "0")
    tr2 = len(dirka)+2
    ws2.merge_cells(f"A{tr2}:C{tr2}"); _t(ws2.cell(tr2,1), "SKUPAJ")
    _t(ws2.cell(tr2,4), f"=SUM(D2:D{tr2-1})", "0")
    _widths(ws2, [14, 36, 18, 20])

    # Sheet 3 — WAU (dejanski čas, samo "wau" dejavnosti)
    ws3 = wb.create_sheet("WAU")
    for c, h in enumerate(["Datum", "Opis", "Ure (dejanski čas)"], 1): _h(ws3.cell(1,c), h)
    for i, r in enumerate(wau):
        _c(ws3.cell(i+2,1), r["datum"], i%2)
        _c(ws3.cell(i+2,2), r["opis"],  i%2)
        _c(ws3.cell(i+2,3), r["ure"],   i%2, "0.00")
    tr3 = len(wau)+2
    ws3.merge_cells(f"A{tr3}:B{tr3}"); _t(ws3.cell(tr3,1), "SKUPAJ")
    _t(ws3.cell(tr3,3), f"=SUM(C2:C{tr3-1})", "0.00")
    _widths(ws3, [14, 32, 22])

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

        rap_h   = sum(r["ure_st"] for r in data["rap_ps"])
        dirka_h = sum(r["ure_st"] for r in data["dirka_za_branje"])
        wau_h   = sum(r["ure"]   for r in data["wau"])
        total   = len(data["rap_ps"]) + len(data["dirka_za_branje"]) + len(data["wau"])

        if total == 0:
            st.error("Ni aktivnosti. Surovo besedilo PDF-ja za debug:")
            st.text_area("", text[:3000], height=200)
        else:
            excel_bytes = build_excel(data)
            st.success("✅ Excel je pripravljen!")
            c1, c2 = st.columns(2)
            c1.markdown(f"**Učitelj:** {data['ucitelj']}")
            c2.markdown(f"**Obdobje:** {data['obdobje']}")
            c3, c4, c5 = st.columns(3)
            c3.metric("RAP / RaP PS",    f"{rap_h} ur",    f"{len(data['rap_ps'])} vnosov")
            c4.metric("Dirka za branje", f"{dirka_h} ur",  f"{len(data['dirka_za_branje'])} vnosov")
            c5.metric("WAU",             f"{wau_h:.2f} ur", f"{len(data['wau'])} vnosov")

            with st.expander("🔎 Predogled prepoznanih vrstic"):
                t1, t2, t3 = st.tabs(["RAP / RaP PS", "Dirka za branje", "WAU"])
                t1.dataframe(data["rap_ps"], use_container_width=True, hide_index=True)
                t2.dataframe(data["dirka_za_branje"], use_container_width=True, hide_index=True)
                t3.dataframe(data["wau"], use_container_width=True, hide_index=True)

            with st.expander("🧪 Surovo besedilo PDF-ja"):
                st.text_area("", text[:8000], height=240)

            safe = re.sub(r'[^\w]', '_', data["ucitelj"])
            st.download_button(
                "⬇️ Prenesi Excel", excel_bytes,
                file_name=f"porocilo_{safe}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

st.markdown("---")
st.caption("Šmarje-Sap · Poročilo opravljenih ur · v2.2")
