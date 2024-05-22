from flask import Flask, render_template, request
import PyPDF2
import re
import os
import openpyxl

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/procesar', methods=['POST'])
def procesar_facturas():
    directorio = request.files.getlist('directorio')
    
    # Crear un libro de Excel para almacenar los resultados
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Facturas"
    ws.append(["Proveedor", "Factura Número", "Cliente", "Items", "Vencimiento", "Total"])
    
    for pdf_file in directorio:
        pdf_filename = pdf_file.filename
        pdf_path = f"uploads/{pdf_filename}"
        pdf_file.save(pdf_path)
        
        # Procesar la factura PDF y guardar los resultados en el archivo Excel
        process_invoice_pdf(pdf_path, ws)

    # Guardar el archivo Excel con los resultados
    output_excel = "facturas.xlsx"
    wb.save(output_excel)
    
    return f"Facturas procesadas con éxito. Resultados guardados en {output_excel}"

def process_invoice_pdf(pdf_file_path, worksheet):
    # Función para procesar una factura PDF y agregar los resultados a la hoja de cálculo
    with open(pdf_file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ''

        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()

        # Patrones de regex para extraer información de la factura
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

        # Agregar los resultados a la hoja de cálculo
        items_str = "\n".join([f"{item[0]}: {item[1]}  = {item[3]}" for item in items])
        worksheet.append([issued_by, invoice_number, bill_to, items_str, due_date, total])

if __name__ == '__main__':
    app.run(debug=True)
