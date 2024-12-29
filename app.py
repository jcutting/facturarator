import streamlit as st

st.set_page_config(layout="wide")

import xml.etree.ElementTree as ET
import csv
import io
from io import BytesIO
import pandas as pd

def validate_xml_data(data):
    required_fields = ['UUID_Last_12', 'RFC_Emisor', 'Total_Impuestos', 'Total_Comprobante']
    missing_fields = []
    for field in required_fields:
        if not data[required_fields.index(field)]:
            missing_fields.append(field)
    return missing_fields

def parse_xml_file(xml_content):
    namespaces = {
        'cfdi': 'http://www.sat.gob.mx/cfd/4',
        'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'
    }
    
    try:
        tree = ET.parse(BytesIO(xml_content))
        root = tree.getroot()
        
        # Extract UUID
        tfd = root.find('.//tfd:TimbreFiscalDigital', namespaces)
        uuid = tfd.get('UUID')[-12:] if tfd is not None else ''
        
        # Extract Emisor details
        emisor = root.find('./cfdi:Emisor', namespaces)
        rfc = emisor.get('Rfc') if emisor is not None else ''
        nombre = emisor.get('Nombre') if emisor is not None else ''
        
        # Extract TotalImpuestosTrasladados
        impuestos = root.find('./cfdi:Impuestos', namespaces)
        total_impuestos = impuestos.get('TotalImpuestosTrasladados') if impuestos is not None else '0'
        
        # Extract Total and Currency
        total = root.get('Total', '0')
        currency = root.get('Moneda', '')
        
        # Check for gasoline terms
        xml_string = ET.tostring(root, encoding='unicode')
        invoice_type = 'Gasoline' if 'Gasoline' in xml_string or 'Gasolina' in xml_string else 'Miscellaneous'
        
        # Return array matching headers order: [Nombre, UUID_Last_12, RFC_Emisor, Total_Impuestos, Total_Comprobante, Type, Currency]
        return [nombre, uuid, rfc, total_impuestos, total, invoice_type, currency]
    
    except Exception as e:
        return ['', '', '', '', '', '', '']  # 7 empty values to match expected columns

def main():
    st.title("XML Invoice Processor")
    
    uploaded_files = st.file_uploader("Upload XML files", type=['xml'], accept_multiple_files=True)
    
    if uploaded_files:
        successful_files = []
        error_files = []
        all_data = []
        headers = ['No.VDR', 'Nombre', 'UUID_Last_12', 'RFC_Emisor', 'Total_Impuestos', 
                  'Total_Comprobante', 'Type', 'Currency']
        
        # Process each file
        for idx, uploaded_file in enumerate(uploaded_files, 1):
            xml_content = uploaded_file.read()
            data = parse_xml_file(xml_content)
            if len(data) == 7:  # Verify we have correct number of columns
                successful_files.append(uploaded_file.name)
                all_data.append([idx] + data)  # Add index as No.VDR
            else:
                error_files.append({
                    'filename': uploaded_file.name,
                    'missing_fields': 'Invalid data structure'
                })
        
        if successful_files:
            df = pd.DataFrame(all_data, columns=headers)
            st.success(f"Successfully processed {len(successful_files)} files")
            st.write("Processed files:")
            st.write(successful_files)
            st.write("Processed Data:")
            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True
            )
            
            csv_buffer = io.StringIO()
            df.to_csv(csv_buffer, index=False)
            csv_str = csv_buffer.getvalue()
            
            st.download_button(
                label="Download CSV",
                data=csv_str,
                file_name="facturas_summary.csv",
                mime="text/csv"
            )
        
        if error_files:
            st.error("Some files had errors:")
            error_df = pd.DataFrame(error_files)
            st.dataframe(error_df)
if __name__ == "__main__":
    main()