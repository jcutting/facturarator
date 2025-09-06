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
        retur
