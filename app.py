import io
import json
import os
import zipfile
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
