import io, json, zipfile
import streamlit as st
import pandas as pd
from datetime import date
from xml.etree import ElementTree as ET

st.set_page_config(page_title="Personal IVA Builder", layout="wide")

def parse_cfdi(xml_bytes):
    ns = {
        "cfdi33": "http://www.sat.gobmx/cfd/3",
        "cfdi33_alt": "http://www.sat.gob.mx/cfd/3",
        "cfdi40": "http://www.sat.gob.mx/cfd/4",
        "cfdi40_alt": "http://www.sat.gobmx/cfd/4",
        "tfd":    "http://www.sat.gob.mx/TimbreFiscalDigital",
    }
    root = ET.fromstring(xml_bytes)
    tag = root.tag
    ckey = "cfdi40" if ("cfd/4" in tag or "cfd/4" in json.dumps(root.attrib)) else "cfdi33"
    cns = ns.get(ckey, ns["cfdi40"])

    def f(path):  return root.find(path, {"cfdi": cns, "tfd": ns["tfd"]})
    def fa(path): return root.findall(path, {"cfdi": cns, "tfd": ns["tfd"]})

    uuid   = (f(".//tfd:TimbreFiscalDigital") or {}).get("UUID","") if f(".//tfd:TimbreFiscalDigital") is not None else ""
    fecha  = (root.get("Fecha","") or "")[:10]
    moneda = root.get("Moneda","") or ""
    tc     = root.get("TipoCambio","") or ""
    subtotal = root.get("SubTotal","") or ""
    total    = root.get("Total","") or ""
    metodo   = root.get("MetodoPago","") or ""
    forma    = root.get("FormaPago","") or ""
    serie    = root.get("Serie","") or ""
    folio    = root.get("Folio","") or ""

    em = f("./cfdi:Emisor"); re = f("./cfdi:Receptor")
    em_nom = em.get("Nombre","") if em is not None else ""
    em_rfc = em.get("Rfc","")    if em is not None else ""
    re_nom = re.get("Nombre","") if re is not None else ""
    re_rfc = re.get("Rfc","")    if re is not None else ""
    uso    = re.get("UsoCFDI","") if re is not None else ""

    iva_sum = 0.0
    if f("./cfdi:Impuestos") is not None:
        for t in fa(".//cfdi:Traslados/cfdi:Traslado"):
            if t.get("Impuesto") in ("002","2"):
                try: iva_sum += float(t.get("Importe","0") or "0")
                except: pass

    factura_number = f"{serie}-{folio}".strip("-") if (serie or folio) else (uuid[-8:] if uuid else "")

    return {
        "FacturaNumber": factura_number,
        "FolioFiscal_UUID": uuid,
        "Fecha": fecha,
        "Proveedor_Nombre": em_nom,
        "Proveedor_RFC": em_rfc,
        "Receptor_Nombre": re_nom,
        "Receptor_RFC": re_rfc,
        "Moneda": moneda,
        "TipoCambio": tc,
        "Subtotal": subtotal,
        "IVA_Trasladado": iva_sum,
        "Total": total,
        "MetodoPago": metodo,
        "FormaPago": forma,
        "UsoCFDI": uso,
        "Categoria": "Miscellaneous",
    }

def build_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        cols = ["FacturaNumber","FolioFiscal_UUID","Fecha","Proveedor_Nombre","Proveedor_RFC",
                "Receptor_Nombre","Receptor_RFC","Moneda","TipoCambio","Subtotal","IVA_Trasladado",
                "Total","MetodoPago","FormaPago","UsoCFDI","Categoria","PDF_FileName","XML_FileName","Notas"]
        for c in cols:
            if c not in df.columns: df[c] = ""
        df = df[cols]
        df.to_excel(writer, index=False, sheet_name="Facturas")
        wb = writer.book; ws = writer.sheets["Facturas"]

        settings = pd.DataFrame({"Monedas":["MXN","USD"], "Categorias":["Miscellaneous","Gasoline"]})
        settings.to_excel(writer, index=False, sheet_name="Settings")
        wb.define_name("Monedas_List", "=Settings!$A$2:$A$3")
        wb.define_name("Categorias_List", "=Settings!$B$2:$B$3")

        txt = wb.add_format({}); money = wb.add_format({"num_format": "#,##0.00"}); datefmt = wb.add_format({"num_format": "yyyy-mm-dd"})
        widths = [18,40,14,28,16,28,16,8,10,14,14,14,12,12,10,14,28,28,30]
        for i,w in enumerate(widths): ws.set_column(i, i, w, txt)
        ws.set_column(9, 11, 14, money); ws.set_column(2, 2, 14, datefmt)
        last = max(1000, len(df)+50)
        ws.data_validation(1, 7, last, 7,  {"validate":"list","source":"=Monedas_List"})
        ws.data_validation(1, 15, last, 15, {"validate":"list","source":"=Categorias_List"})
        ws.data_validation(1, 1, last, 1,  {
            "validate":"length","criteria":"equal to","value":36,
            "input_title":"Folio Fiscal (UUID)","input_message":"Debe tener exactamente 36 caracteres (incluye guiones).",
            "error_title":"Longitud inválida","error_message":"El UUID debe tener 36 caracteres."
        })

        notes = [
            "Checklist (Personal IVA – Digital Submission)",
            "1) Review 'Facturas' and adjust dropdowns (Moneda, Categoria) as needed.",
            "2) Verify UUID is 36 chars and matches the CFDI Timbre.",
            "3) Name each PDF with the FacturaNumber (column A).",
            "4) Build ZIP of PDFs (≤ 20 MB).",
            "5) Include SRE Carnet PDF in the email (warning only).",
            f"Generated: {date.today().isoformat()}",
        ]
        pd.DataFrame({"Notes": notes}).to_excel(writer, index=False, sheet_name="Instructions")
        writer.sheets["Instructions"].set_column(0,0,120)
    buf.seek(0)
    return buf.getvalue()

st.title("Personal IVA – Excel & ZIP Builder")

xml_up = st.file_uploader("Upload CFDI XML files", type=["xml"], accept_multiple_files=True)
pdf_up = st.file_uploader("Upload matching PDF facturas", type=["pdf"], accept_multiple_files=True)
carnet_up = st.file_uploader("Upload SRE Carnet (optional – warning only)", type=["pdf"])

rows = []
if xml_up:
    for f in xml_up:
        try:
            row = parse_cfdi(f.read())
            row["XML_FileName"] = f.name
            rows.append(row)
        except Exception as e:
            rows.append({"XML_FileName": f.name, "Notas": f"Parse error: {e}"})
    df = pd.DataFrame(rows)
    # Map PDFs by name contains FacturaNumber or UUID prefix
    pdf_names = [p.name for p in (pdf_up or [])]
    def guess_pdf(r):
        fn = str(r.get("FacturaNumber","") or "")
        uuid = str(r.get("FolioFiscal_UUID","") or "")
        for n in pdf_names:
            if fn and fn.lower() in n.lower(): return n
        for n in pdf_names:
            if uuid and uuid.lower()[:8] in n.lower(): return n
        return ""
    df["PDF_FileName"] = df.apply(guess_pdf, axis=1)
    if "Categoria" not in df.columns: df["Categoria"] = "Miscellaneous"
    df["Notas"] = df.get("Notas","")

    st.subheader("Preview / Edit")
    edited = st.data_editor(
        df,
        num_rows="dynamic",
        column_config={
            "Moneda": st.column_config.SelectboxColumn(options=["MXN","USD"]),
            "Categoria": st.column_config.SelectboxColumn(options=["Miscellaneous","Gasoline"]),
        },
        use_container_width=True,
        height=420
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Generate Excel (2025 format)"):
            xlsx = build_excel(edited.copy())
            st.download_button("Download Facturas_Submission.xlsx", xlsx, file_name="Facturas_Submission.xlsx")
    with c2:
        if st.button("Build ZIP of PDFs (renamed)"):
            if not pdf_up:
                st.error("Please upload the factura PDFs first.")
            else:
                by_name = {p.name: p.getvalue() for p in pdf_up}
                buf = io.BytesIO()
                errors = []
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                    for _, r in edited.iterrows():
                        fact = str(r.get("FacturaNumber","") or "").strip()
                        src  = str(r.get("PDF_FileName","") or "").strip()
                        if not fact:
                            errors.append(f"Missing FacturaNumber for XML={r.get('XML_FileName','')}"); continue
                        if not src or src not in by_name:
                            errors.append(f"Missing PDF for FacturaNumber={fact} (expected '{src}')"); continue
                        z.writestr(f"{fact}.pdf", by_name[src])
                st.download_button("Download Facturas_PDFs.zip", buf.getvalue(), file_name="Facturas_PDFs.zip")
                if errors:
                    st.warning("Some issues:\n- " + "\n- ".join(errors))

    # Carnet warning (no blocking)
    if carnet_up is not None:
        size_mb = len(carnet_up.getvalue()) / (1024*1024)
        if size_mb > 3:
            st.warning(f"Carnet PDF is {size_mb:.2f} MB (> 3 MB). Consider compressing to ≤ 3 MB (portrait).")
        else:
            st.info(f"Carnet PDF size looks OK: {size_mb:.2f} MB (≤ 3 MB).")
else:
    st.info("Upload your XMLs to begin.")
