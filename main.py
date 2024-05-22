import PyPDF2
import re
import os
import openpyxl

def extract_invoice_info(pdf_file_path):
    with open(pdf_file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ''

        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()

        invoice_number_pattern = r'Factura N°:\s*([\d-]+)'
        bill_to_pattern = r'Señor\(es\):\s*(.*?)\n'
        issued_by_pattern = r'^(.*?)\n'
        items_pattern = r'(.*?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})'
        due_date_pattern = r'Fecha de Vencimiento:\s*(\d{2}-\d{2}-\d{4})'
        total_pattern = r'Total:\s*\$?\s*([\d,]+\.\d{2})'

        invoice_number_match = re.search(invoice_number_pattern, text)
        bill_to_match = re.search(bill_to_pattern, text)
        issued_by_match = re.search(issued_by_pattern, text)
        items_matches = re.findall(items_pattern, text)
        due_date_match = re.search(due_date_pattern, text)
        total_match = re.search(total_pattern, text)

        invoice_number = invoice_number_match.group(1) if invoice_number_match else None
        bill_to = bill_to_match.group(1) if bill_to_match else None
        issued_by = issued_by_match.group(1) if issued_by_match else None
        items = items_matches
        due_date = due_date_match.group(1) if due_date_match else None
        total = total_match.group(1) if total_match else None

        return {
            'Proveedor': issued_by,
            'Factura Número': invoice_number,
            'Cliente': bill_to,
            'Items': items,
            'Vencimiento': due_date,
            'Total': total
        }

def process_pdfs_in_folder(folder_path, output_excel):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Facturas"
    ws.append(["Proveedor", "Factura Número", "Cliente", "Items", "Vencimiento", "Total"])

    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            info = extract_invoice_info(pdf_path)

            items_str = "\n".join([f"{item[0]}: {item[1]}  = {item[3]}" for item in info['Items']])

            ws.append([
                info['Proveedor'],
                info['Factura Número'],
                info['Cliente'],
                items_str,
                info['Vencimiento'],
                info['Total']
            ])

    wb.save(output_excel)
    print(f"Datos guardados en {output_excel}")

folder_path = r"C:\xampp\htdocs\pypdffacturas\vallenetfacturaspendientesroca"

output_excel = "facturas.xlsx"

process_pdfs_in_folder(folder_path, output_excel)
