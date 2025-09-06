import io
import os
import re
import json
import zipfile
import unicodedata
from datetime import datetime
from xml.etree import ElementTree as ET

import pandas as pd
import streamlit as st


# ------------------------------
# Helpers
# ------------------------------
def seq_label(n: int, width: int = 2) -> str:
    """Return zero-padded label like '01', '02'."""
    return str(n).zfill(width)

def normalize_name(name: str) -> str:
    """
    Robust filename normalizer for matching:
    - take basename
    - unicode normalize
    - lowercase
    - strip extension if present
    - collapse any non-alphanumeric (including slashes, spaces, colons) to single '-'
    - trim leading/trailing '-'
    """
    if not name:
        return ""
    base = os.path.basename(name)
    # strip extension
    stem, _ = os.path.splitext(base)
    # unicode normalize
    stem = unicodedata.normalize("NFKC", stem)
    stem = stem.lower().strip()
    # collapse non-alnum
    stem = re.sub(r"[^0-9a-z]+", "-", stem)
    stem = stem.strip("-")
    return stem


# ------------------------------
# XML (CFDI 3.3 / 4.0) PARSER
# ------------------------------
def parse_cfdi(xml_bytes):
    """
    Minimal fields needed for the IVA submission grid + Fecha for sorting:
      UUID, RFC_Emisor, Total_Impuestos (IVA trasladado), Total_Comprobante, Currency, Fecha
    """
    ns = {
        "cfdi33": "http://www.sat.gobmx/cfd/3",
        "cfdi33_alt": "http://www.sat.gob.mx/cfd/3",
        "cfdi40": "http://www.sat.gob.mx/cfd/4",
        "cfdi40_alt": "http://www.sat.gobmx/cfd/4",
        "tfd": "http://www.sat.gob.mx/TimbreFiscalDigital",
    }
    root = ET.fromstring(xml_bytes)

    # choose CFDI namespace
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
    rfc_emisor = ""
    em = f("./cfdi:Emisor")
    if em is not None:
        rfc_emisor = em.get("Rfc", "") or em.get("RfcEmisor", "") or ""

    # Currency, totals
    currency = root.get("Moneda", "") or "MXN"
    total = root.get("Total", "") or "0"

    # Fecha (for chronological sorting)
    fecha_raw = (root.get("Fecha", "") or "")[:10]  # 'YYYY-MM-DD'
    try:
        fecha_dt = datetime.strptime(fecha_raw, "%Y-%m-%d")
    except Exception:
        fecha_dt = datetime(1970, 1, 1)  # fallback if missing

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
        "Fecha": fecha_dt,        # internal only (not exported to Excel)
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
    Creates an Excel file "SUBMISSION IVA FORM" laid out like the slide:
    Title row, top info (Requested month, Claimant name, Official email, SSN),
    then the blue header row and the table.

    Expects df columns:
      ['No.VDR','UUID','RFC_Emisor','Total_Impuestos','Total_Comprobante','Type','Currency']
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

        # Ensure expected columns & order
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
            df.loc[0] = ["01", "", "", 0.0, 0.0, "Miscellaneous", "MXN"]

        # Write table rows (No.VDR kept as TEXT with leading zeros)
        start = 5
        for i, row in df.iterrows():
            ws.write(start + i, 0, str(row["No.VDR"]), center)  # keep leading zeros
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
        last = start + max(len(df), 50)
        ws.data_validation(
            start, 5, last, 5, {"validate": "list", "source": ["Miscellaneous", "Gasoline"]}
        )
        ws.data_validation(start, 6, last, 6, {"validate": "list", "source": ["MXN", "USD"]})
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
# STREAMLIT APP (ONE FLAT ZIP)
# ------------------------------
st.set_page_config(page_title="Personal IVA – One-Click Package", layout="wide")
st.title("Personal IVA – One-Click Submission Package (Flat ZIP + Robust Matching)")

st.markdown(
    "Upload **CFDI XML** and **matching PDF facturas** (same filename, different extension).\n\n"
    "This will:\n"
    "1) Parse XML, sort entries **chronologically**, and number rows **01, 02, …**\n"
    "2) Build **SUBMISSION_IVA_FORM.xlsx** (slide layout)\n"
    "3) Create a **single flat ZIP** containing the Excel **and** the renamed PDFs (`01.pdf`, `02.pdf`, …)\n"
    "   \n**Now with robust filename matching** (handles spaces, slashes, punctuation) and UUID fallback."
)

# Claimant details (appear in Excel)
col_a, col_b, col_c, col_d = st.columns([1.2, 1.2, 1, 0.9])
with col_a:
    claimant_name = st.text_input("CLAIMANT NAME", value="", placeholder="First Last")
with col_b:
    official_email = st.text_input("OFFICIAL E-MAIL", value="", placeholder="you@state.gov")
with col_c:
    ssn_last4 = st.text_input("SSN (LAST 4 DIGITS)", value="", max_chars=4)
with col_d:
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
    # Parse all XMLs (collect Fecha for sorting)
    for f in xml_up:
        try:
            row = parse_cfdi(f.read())
            row["XML_FileName"] = f.name
            row["XML_Stem"] = normalize_name(f.name)          # normalized
            row["XML_Stem_raw"] = os.path.splitext(os.path.basename(f.name))[0]
            rows.append(row)
        except Exception as e:
            rows.append(
                {
                    "XML_FileName": f.name,
                    "XML_Stem": normalize_name(f.name),
                    "XML_Stem_raw": os.path.splitext(os.path.basename(f.name))[0],
                    "UUID": "",
                    "RFC_Emisor": "",
                    "Total_Impuestos": 0.0,
                    "Total_Comprobante": 0.0,
                    "Currency": "MXN",
                    "Type": "Miscellaneous",
                    "Fecha": datetime(1970, 1, 1),
                    "Notas": f"Parse error: {e}",
                }
            )

    df = pd.DataFrame(rows)

    # Chronological order by Fecha (ascending)
    df = df.sort_values(by="Fecha", ascending=True).reset_index(drop=True)

    # Assign No.VDR as zero-padded text (01, 02, 03)
    df.insert(0, "No.VDR", [seq_label(i + 1, 2) for i in range(len(df))])

    # Build PDF bytes map by normalized stem
    pdf_map = {}
    pdf_name_lookup = {}   # normalized stem -> original pdf name (for messages)
    if pdf_up:
        for p in pdf_up:
            nstem = normalize_name(p.name)
            pdf_map[nstem] = p.getvalue()
            pdf_name_lookup[nstem] = os.path.basename(p.name)

    # Match PDFs by normalized stem; fallback by UUID substring
    def pdf_name_for_row(r):
        nstem = r.get("XML_Stem", "")
        uuid = str(r.get("UUID", "") or "").lower()
        if nstem in pdf_map:
            return pdf_name_lookup[nstem]
        # fallback: search any PDF name containing UUID (or first 8 chars)
        if uuid:
            short = uuid[:8]
            for key, orig_name in pdf_name_lookup.items():
                if uuid in key or short in key:
                    return orig_name
        return ""

    df["PDF_FileName"] = df.apply(pdf_name_for_row, axis=1)

    # Preview / Edit
    st.subheader("Preview / Edit (chronological)")
    show_cols = [
        "No.VDR",
        "UUID",
        "RFC_Emisor",
        "Total_Impuestos",
        "Total_Comprobante",
        "Type",
        "Currency",
        "Fecha",
        "XML_FileName",
        "PDF_FileName",
    ]
    edited = st.data_editor(
        df[show_cols],
        num_rows="dynamic",
        column_config={
            "Type": st.column_config.SelectboxColumn(options=["Miscellaneous", "Gasoline"]),
            "Currency": st.column_config.SelectboxColumn(options=["MXN", "USD"]),
        },
        use_container_width=True,
        height=480,
    )

    # One-click: build Excel + flat ZIP (Excel + PDFs at root)
    if st.button("Build Submission Package (flat ZIP)"):
        # Re-sort and re-label to ensure consistency with any edits
        tmp = edited.copy()
        if "Fecha" in tmp.columns:
            tmp = tmp.sort_values(by="Fecha", ascending=True).reset_index(drop=True)
        tmp["No.VDR"] = [seq_label(i + 1, 2) for i in range(len(tmp))]

        # Validate & gather warnings
        warnings = []
        bad_uuid_rows = [str(r["No.VDR"]) for _, r in tmp.iterrows() if len(str(r.get("UUID",""))) != 36]
        if bad_uuid_rows:
            warnings.append("UUID not 36 chars for rows: " + ", ".join(bad_uuid_rows))

        # Build a fresh normalized pdf map (in case names were re-uploaded)
        pdf_map = {}
        pdf_name_lookup = {}
        if pdf_up:
            for p in pdf_up:
                nstem = normalize_name(p.name)
                pdf_map[nstem] = p.getvalue()
                pdf_name_lookup[nstem] = os.path.basename(p.name)

        # Build Excel (export only required columns)
        export_df = tmp[
            ["No.VDR", "UUID", "RFC_Emisor", "Total_Impuestos", "Total_Comprobante", "Type", "Currency"]
        ].copy()
        xlsx_bytes = build_submission_excel_from_df(
            export_df,
            claimant_name=claimant_name.strip(),
            official_email=official_email.strip(),
            ssn_last4=ssn_last4.strip(),
            requested_month_text=requested_month,
        )

        # Build ONE flat package: Excel + PDFs at root (01.pdf, 02.pdf, ...)
        missing_pdfs_detail = []
        outer = io.BytesIO()
        with zipfile.ZipFile(outer, "w", zipfile.ZIP_DEFLATED) as z:
            # Excel at root
            z.writestr("SUBMISSION_IVA_FORM.xlsx", xlsx_bytes)
            # PDFs at root, renamed by row number
            for _, r in tmp.iterrows():
                rownum = str(r.get("No.VDR", "") or "").strip()
                # Use normalized XML stem for lookup; fallback to UUID match
                xml_stem_raw = os.path.splitext(os.path.basename(str(r.get("XML_FileName",""))))[0]
                nstem = normalize_name(xml_stem_raw)
                uuid = str(r.get("UUID","") or "").lower()
                if rownum:
                    if nstem in pdf_map:
                        z.writestr(f"{rownum}.pdf", pdf_map[nstem])
                    else:
                        # fallback by UUID substring
                        placed = False
                        if uuid:
                            short = uuid[:8]
                            for key, data in pdf_map.items():
                                if uuid in key or short in key:
                                    z.writestr(f"{rownum}.pdf", data)
                                    placed = True
                                    break
                        if not placed:
                            missing_pdfs_detail.append(f"{rownum} (expected stem ~ '{nstem}')")

            # Optional readme
            readme = (
                "This package contains:\n"
                "- SUBMISSION_IVA_FORM.xlsx (Excel to email)\n"
                "- PDFs named 01.pdf, 02.pdf, ... (rename rule = No. VDR)\n"
                "\n"
                "If any warnings were shown in the app, please resolve and rebuild.\n"
            )
            z.writestr("README.txt", readme)
        outer.seek(0)

        st.download_button(
            "Download Submission_Package.zip",
            outer.getvalue(),
            file_name="Submission_Package.zip",
        )

        # Show warnings (non-blocking)
        if carnet_up is not None:
            size_mb = len(carnet_up.getvalue()) / (1024 * 1024)
            if size_mb > 3:
                warnings.append(f"Carnet PDF is {size_mb:.2f} MB (>3 MB). Consider compressing to ≤ 3 MB.")

        if missing_pdfs_detail:
            warnings.append("Missing matching PDF for rows: " + ", ".join(missing_pdfs_detail))

        if warnings:
            st.warning("Warnings:\n- " + "\n- ".join(warnings))

else:
    st.info("Upload your CFDI XML files to begin.")
