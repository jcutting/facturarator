import io
import json
import zipfile
from datetime import datetime
from xml.etree import ElementTree as ET

import pandas as pd
import streamlit as st


# ------------------------------
# XML (CFDI 3.3 / 4.0) PARSER
# ------------------------------
def parse_cfdi(xml_bytes):
    """
    Returns the minimal fields needed for the IVA submission grid:
      UUID, RFC_Emisor, Total_Impuestos (IVA trasladado), Total_Comprobante, Currency
    """
    ns = {
        "cfdi33": "http://www.sat.gobmx/cfd/3",
        "cfdi33_alt": "http://www.sat.gob.mx/cfd/3",
        "cfdi40": "http://www.sat.gob.mx/cfd/4",
        "cfdi40_alt": "http://www.sat.gobmx/cfd/4",
        "tfd": "http://www.sat.gob.mx/TimbreFiscalDigital",
    }
    root = ET.fromstring(xml_bytes)

    # pick CFDI namespace
    tag = root.tag
    ckey = "cfdi40" if ("cfd/4" in tag or "cfd/4" in json.dumps(root.attrib)) else "cfdi33"
    cns = ns.get(ckey, ns["cfdi40"])

    def f(path):
        return root.find(path, {"cfdi": cns, "tfd": ns["tfd"]})

    def fa(path):
        return root.findall(path, {"cfdi": cns, "tfd": ns["tfd"]})

    # UUID
    uuid = ""
    tfd = f(".//tfd:TimbreFiscalDigital")
    if tfd is not None:
        uuid = tfd.get("UUID", "") or ""

    # RFC Emisor
    em = f("./cfdi:Emisor")
    rfc_emisor = ""
    if em is not None:
        rfc_emisor = em.get("Rfc", "") or em.get("RfcEmisor", "") or ""

    # Currency and totals
    currency = root.get("Moneda", "") or "MXN"
    total = root.get("Total", "") or "0"

    # IVA (Traslados Impuesto=002)
    iva_sum = 0.0
    for t in fa(".//cfdi:Traslados/cfdi:Traslado"):
        if t.get("Impuesto") in ("002", "2"):
            try:
                iva_sum += float(t.get("Importe", "0") or "0")
            except Exception:
                pass

    return {
        "UUID": uuid,
        "RFC_Emisor": rfc_emisor,
        "Total_Impuestos": iva_sum,
        "Total_Comprobante": total,
        "Currency": currency,
        "Type": "Miscellaneous",  # default; user can change in UI
    }


# ------------------------------
# EXCEL BUILDER (MATCHES SLIDE)
# ------------------------------
def build_submission_excel_from_df(
    df: pd.DataFrame,
    claimant_name: str,
    official_email: str,
    ssn_last4: str,
    requested_month_text: str,
) -> bytes:
    """
    Creates an Excel file with a single sheet "SUBMISSION IVA FORM" laid out like the slide:
    Title row, top info row (Requested month, Claimant name, Official email, SSN),
    then the blue header row and the table.

    Columns expected in df: ['No.VDR','UUID','RFC_Emisor','Total_Impuestos',
                             'Total_Comprobante','Type','Currency']
    """
    import xlsxwriter

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb: xlsxwriter.Workbook = writer.book
        ws = wb.add_worksheet("SUBMISSION IVA FORM")
        writer.sheets["SUBMISSION IVA FORM"] = ws

        # Formats
        title_fmt = wb.add_format(
            {"bold": True, "align": "center", "valign": "vcenter", "font_size": 14}
        )
        head_lbl = wb.add_format({"bold": True, "font_color": "green"})
        head_val = wb.add_format({"bold": True})
        header_blue = wb.add_format(
            {
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#4F81BD",
                "font_color": "white",
                "border": 1,
            }
        )
        cell = wb.add_format({"border": 1})
        money = wb.add_format({"num_format": "#,##0.00", "border": 1})
        center = wb.add_format({"align": "center", "border": 1})

        # Column widths (A..G)
        ws.set_column("A:A", 10)  # No. VDR
        ws.set_column("B:B", 44)  # UUID (Factura Number per slide)
        ws.set_column("C:C", 18)  # RFC
        ws.set_column("D:D", 14)  # VAT
        ws.set_column("E:E", 14)  # TOTAL
        ws.set_column("F:F", 18)  # Type of expense
        ws.set_column("G:G", 10)  # Currency

        # Title
        ws.merge_range("A1:G1", "SUBMISSION IVA FORM", title_fmt)

        # Top info rows (exact labels from slide)
        ws.write("A2", "REQUESTED MONTH:", head_lbl)
        ws.write("B2", requested_month_text, head_val)

        ws.write("C2", "", head_lbl)  # spacer cell in slide
        ws.write("D2", "OFFICIAL E-MAIL:", head_lbl)
        ws.write("E2", official_email or "", head_val)

        ws.write("F2", "SSN (LAST 4 DIGITS):", head_lbl)
        ws.write("G2", (ssn_last4 or ""), head_val)

        ws.write("A3", "CLAIMANT NAME:", head_lbl)
        ws.write("B3", claimant_name or "", head_val)

        # Blue header row
        ws.write("A5", "No. VDR", header_blue)
        ws.write("B5", "FACTURA NUMBER (COMPLETE FOLIO FISCAL 36 DIGITS)", header_blue)
        ws.write("C5", "R.F.C. FROM VENDOR", header_blue)
        ws.write("D5", "VAT AMOUNT", header_blue)
        ws.write("E5", "TOTAL AMOUNT", header_blue)
        ws.write("F5", "TYPE OF EXPENSE", header_blue)
        ws.write("G5", "CURRENCY", header_blue)

        # Ensure expected columns exist and order them
        needed = [
            "No.VDR",
            "UUID",
            "RFC_Emisor",
            "Total_Impuestos",
            "Total_Comprobante",
            "Type",
            "Currency",
        ]
        for c in needed:
            if c not in df.columns:
                df[c] = ""

        df = df[needed].copy()
        if df.empty:
            # Write one empty row so the file isn't "blank"
            df.loc[0] = [1, "", "", 0.0, 0.0, "Miscellaneous", "MXN"]

        # Auto-fill No.VDR if missing / blank
        if df["No.VDR"].isna().any() or (df["No.VDR"] == "").any():
            df["No.VDR"] = range(1, len(df) + 1)

        # Write table rows
        start = 5  # row index where data body starts (0-based)
        for i, row in df.iterrows():
            ws.write_number(start + i, 0, int(row["No.VDR"]), center)
            ws.write(start + i, 1, str(row["UUID"]), cell)
            ws.write(start + i, 2, str(row["RFC_Emisor"]), cell)

            # numeric VAT/TOTAL
            try:
                vat = float(row["Total_Impuestos"] or 0)
            except Exception:
                vat = 0.0
            try:
                tot = float(row["Total_Comprobante"] or 0)
            except Exception:
                tot = 0.0

            ws.write_number(start + i, 3, vat, money)
            ws.write_number(start + i, 4, tot, money)
            ws.write(start + i, 5, str(row["Type"] or "Miscellaneous"), cell)
            ws.write(start + i, 6, str(row["Currency"] or "MXN"), center)

        # Validations on body range
        last = start + max(len(df), 50)  # give some extra editable rows
        # Type dropdown
        ws.data_validation(
            start, 5, last, 5, {"validate": "list", "source": ["Miscellaneous", "Gasoline"]}
        )
        # Currency dropdown
        ws.data_validation(start, 6, last, 6, {"validate": "list", "source": ["MXN", "USD"]})
        # UUID length must be 36 characters
        ws.data_validation(
            start,
            1,
            last,
            1,
            {
                "validate": "length",
                "criteria": "equal to",
                "value": 36,
                "input_title": "Folio Fiscal (UUID)",
                "input_message": "Debe tener exactamente 36 caracteres (incluye guiones).",
                "error_title": "Longitud inválida",
                "error_message": "El UUID debe tener 36 caracteres.",
            },
        )

    buf.seek(0)
    return buf.getvalue()


# ------------------------------
# STREAMLIT APP
# ------------------------------
st.set_page_config(page_title="Personal IVA – Excel & ZIP Builder", layout="wide")
st.title("Personal IVA – Excel & ZIP Builder")

st.markdown(
    "Upload your **CFDI XML** and **matching PDF facturas**.\n\n"
    "- The app parses values from XML and fills the grid automatically.\n"
    "- PDFs will be **renamed to the UUID** (Factura Number) when building the ZIP.\n"
    "- Provide claimant details below; they appear in the Excel header."
)

# Claimant details (appear in Excel)
col_a, col_b, col_c, col_d = st.columns([1.2, 1.2, 1, 0.8])
with col_a:
    claimant_name = st.text_input("CLAIMANT NAME", value="", placeholder="First Last")
with col_b:
    official_email = st.text_input("OFFICIAL E-MAIL", value="", placeholder="you@state.gov")
with col_c:
    ssn_last4 = st.text_input("SSN (LAST 4 DIGITS)", value="", max_chars=4)
with col_d:
    # Requested month auto: current month + year
    now = datetime.now()
    requested_month = now.strftime("%B %Y")
    st.write("**Requested month**")
    st.info(requested_month)

# Uploaders
xml_up = st.file_uploader("Upload CFDI XML files", type=["xml"], accept_multiple_files=True)
pdf_up = st.file_uploader("Upload matching PDF facturas", type=["pdf"], accept_multiple_files=True)
carnet_up = st.file_uploader("Upload SRE Carnet (optional – warning only)", type=["pdf"])

rows = []
if xml_up:
    # Parse all XMLs
    for f in xml_up:
        try:
            row = parse_cfdi(f.read())
            row["XML_FileName"] = f.name
            rows.append(row)
        except Exception as e:
            rows.append(
                {
                    "XML_FileName": f.name,
                    "UUID": "",
                    "RFC_Emisor": "",
                    "Total_Impuestos": 0.0,
                    "Total_Comprobante": 0.0,
                    "Type": "Miscellaneous",
                    "Currency": "MXN",
                    "Notas": f"Parse error: {e}",
                }
            )

    df = pd.DataFrame(rows)

    # Insert No.VDR as 1..N
    df.insert(0, "No.VDR", range(1, len(df) + 1))

    # Best-effort: try to auto-match PDFs by UUID substring
    pdf_names = [p.name for p in (pdf_up or [])]

    def guess_pdf_for_row(r):
        uuid = str(r.get("UUID", "") or "").lower()
        for n in pdf_names:
            nm = n.lower()
            if uuid and (uuid in nm or uuid[:8] in nm):
                return n
        return ""

    df["PDF_FileName"] = df.apply(guess_pdf_for_row, axis=1)

    # Show editable grid (Type + Currency as dropdowns)
    st.subheader("Preview / Edit")
    edited = st.data_editor(
        df[
            [
                "No.VDR",
                "UUID",
                "RFC_Emisor",
                "Total_Impuestos",
                "Total_Comprobante",
                "Type",
                "Currency",
                "XML_FileName",
                "PDF_FileName",
            ]
        ],
        num_rows="dynamic",
        column_config={
            "Type": st.column_config.SelectboxColumn(options=["Miscellaneous", "Gasoline"]),
            "Currency": st.column_config.SelectboxColumn(options=["MXN", "USD"]),
        },
        use_container_width=True,
        height=460,
    )

    st.caption(
        "Note: the Excel **only** contains the columns shown in the slide. "
        "Other columns here (like XML/PDF file names) are for your convenience."
    )

    c1, c2 = st.columns(2)

    with c1:
        if st.button("Generate Excel (Submission Template)"):
            xlsx_bytes = build_submission_excel_from_df(
                edited.copy(),
                claimant_name=claimant_name.strip(),
                official_email=official_email.strip(),
                ssn_last4=ssn_last4.strip(),
                requested_month_text=requested_month,
            )
            st.download_button(
                "Download SUBMISSION_IVA_FORM.xlsx",
                xlsx_bytes,
                file_name="SUBMISSION_IVA_FORM.xlsx",
            )

    with c2:
        if st.button("Build ZIP of PDFs (renamed to UUID)"):
            if not pdf_up:
                st.error("Please upload the factura PDFs first.")
            else:
                # map uploads by name
                by_name = {p.name: p.getvalue() for p in pdf_up}
                buf = io.BytesIO()
                errors = []
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                    for _, r in edited.iterrows():
                        uuid = str(r.get("UUID", "") or "").strip()
                        src = str(r.get("PDF_FileName", "") or "").strip()
                        if not uuid:
                            errors.append(
                                f"Missing UUID for row No.VDR={r.get('No.VDR','?')} (cannot rename)."
                            )
                            continue
                        if not src or src not in by_name:
                            errors.append(
                                f"Missing PDF for UUID={uuid} (expected file name '{src}')."
                            )
                            continue
                        z.writestr(f"{uuid}.pdf", by_name[src])
                st.download_button(
                    "Download Facturas_PDFs.zip", buf.getvalue(), file_name="Facturas_PDFs.zip"
                )
                if errors:
                    st.warning("Some issues:\n- " + "\n- ".join(errors))

    # Carnet warning only (no blocking)
    if carnet_up is not None:
        size_mb = len(carnet_up.getvalue()) / (1024 * 1024)
        if size_mb > 3:
            st.warning(
                f"Carnet PDF is {size_mb:.2f} MB (>3 MB). Consider compressing to ≤ 3 MB (portrait)."
            )
        else:
            st.info(f"Carnet PDF size looks OK: {size_mb:.2f} MB (≤ 3 MB).")

else:
    st.info("Upload your CFDI XML files to begin.")
