import xml.etree.ElementTree as ET
import csv
import os
import glob

def parse_xml_file(file_path):
    # Define XML namespace mapping
    namespaces = {
        'cfdi': 'http://www.sat.gob.mx/cfd/4',
        'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'
    }
    
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
        
        # Extract UUID (last 12 characters)
        tfd = root.find('.//tfd:TimbreFiscalDigital', namespaces)
        uuid = tfd.get('UUID')[-12:] if tfd is not None else ''
        
        # Extract RFC
        emisor = root.find('.//cfdi:Emisor', namespaces)
        rfc = emisor.get('Rfc') if emisor is not None else ''
        print(emisor)
        
        # Extract TotalImpuestosTrasladados
        impuestos = root.find('./cfdi:Impuestos', namespaces)
        total_impuestos = impuestos.get('TotalImpuestosTrasladados') if impuestos is not None else '0'
        print("Impuestos node:", ET.tostring(impuestos) if impuestos is not None else "Not found")
        print(total_impuestos)
        
        # Extract Total amount
        total = root.get('Total', '0')
        
        return [uuid, rfc, total_impuestos, total]
    
    except Exception as e:
        print(f"Error processing file {file_path}: {str(e)}")
        return ['', '', '', '']

def main():
    # Output CSV file
    output_file = 'facturas_summary.csv'
    headers = ['UUID_Last_12', 'RFC_Emisor', 'Total_Impuestos', 'Total_Comprobante']
    
    # Get all XML files in current directory
    xml_files = glob.glob('*.xml')
    
    with open(output_file, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(headers)
        
        for xml_file in xml_files:
            data = parse_xml_file(xml_file)
            writer.writerow(data)
            
    print(f"Processing complete. Results saved to {output_file}")

if __name__ == "__main__":
    main()