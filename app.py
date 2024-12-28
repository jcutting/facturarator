import streamlit as st
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
        
        # Extract UUID (last 12 characters)
        tfd = root.find('.//tfd:TimbreFiscalDigital', namespaces)
        uuid = tfd.get('UUID')[-12:] if tfd is not None else ''
        
        # Extract RFC
        emisor = root.find('./cfdi:Emisor', namespaces)
        rfc = emisor.get('Rfc') if emisor is not None else ''
        
        # Extract TotalImpuestosTrasladados
        impuestos = root.find('./cfdi:Impuestos', namespaces)
        total_impuestos = impuestos.get('TotalImpuestosTrasladados') if impuestos is not None else '0'
        
        # Extract Total amount
        total = root.get('Total', '0')
        
        return [uuid, rfc, total_impuestos, total]
    
    except Exception as e:
        return ['', '', '', '']

def main():
    st.title("XML Invoice Processor")
    
    # File uploader
    uploaded_files = st.file_uploader("Upload XML files", type=['xml'], accept_multiple_files=True)
    
    if uploaded_files:
        # Initialize results storage
        successful_files = []
        error_files = []
        all_data = []
        headers = ['UUID_Last_12', 'RFC_Emisor', 'Total_Impuestos', 'Total_Comprobante']
        
        # Process each file
        for uploaded_file in uploaded_files:
            xml_content = uploaded_file.read()
            data = parse_xml_file(xml_content)
            missing_fields = validate_xml_data(data)
            
            if missing_fields:
                error_files.append({
                    'filename': uploaded_file.name,
                    'missing_fields': ', '.join(missing_fields)
                })
            else:
                successful_files.append(uploaded_file.name)
                all_data.append(data)
        
        # Display results
        if successful_files:
            st.success(f"Successfully processed {len(successful_files)} files")
            st.write("Processed files:")
            st.write(successful_files)
            
            # Create DataFrame and display
            df = pd.DataFrame(all_data, columns=headers)
            st.write("Processed Data:")
            st.dataframe(df)
            
            # Create download button for CSV
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