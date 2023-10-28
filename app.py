import subprocess
import pdfplumber
import pandas as pd
import tabula
from pdfminer.high_level import extract_text
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, jsonify
import os
from fpdf import FPDF
from img2pdf import convert
from openpyxl import load_workbook
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
import csv
from pdf2image import convert_from_path
# Naming current module
app = Flask(__name__)

# Define upload and converted PDF directories
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted_pdfs'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER

# Render the home page for file conversion
@app.route('/')
def index():
    return render_template('index.html')

# Function to convert images to PDF
def convert_image_to_pdf(image_path, output_path):
    with open(output_path, 'wb') as pdf_file, open(image_path, 'rb') as image_file:
        pdf_file.write(convert(image_file.read()))

# function to convert document types to PDF
def convert_doc_to_pdf(document_path, output_path):
    try:
        #Using libreoffice to convert the document to PDF
        subprocess_args = ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(output_path), document_path]
        subprocess.run(subprocess_args, check=True)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Conversion failed: {e}")
        return False

# Function to convert PPTX to PDF
def convert_ppt_to_pdf(ppt_path, output_path):
    # Using libre office for converting ppt to PPF
    subprocess_args = ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(output_path), ppt_path]
    subprocess.run(subprocess_args)

# Function to convert CSV to PDF
def convert_csv_to_pdf(csv_path, pdf_path):
    #Reading the CSV file
    data = []
    with open(csv_path, 'r') as csvfile:
        csvreader = csv.reader(csvfile)
        for row in csvreader:
            data.append(row)
    #Creating a pdf template using reportlab
    doc = SimpleDocTemplate(pdf_path, pagesize=letter)
    #Genearing table with data
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), (0.9, 0.9, 0.9)),
        ('TEXTCOLOR', (0, 0), (-1, 0), (0, 0, 0)),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), (0.85, 0.85, 0.85)),
        ('GRID', (0, 0), (-1, -1), 1, (0, 0, 0)),
    ]))
    doc.build([table])

# Function to convert Excel to PDF
def convert_excel_to_pdf(excel_path, output_path):
    #Converting using FPDF
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    doc = SimpleDocTemplate(output_path, pagesize=letter, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=18)
    elements = []
    workbook = load_workbook(excel_path)
    #Reading Sheets from Excel
    for sheet in workbook.sheetnames:
        data = []
        for row in workbook[sheet].iter_rows(values_only=True):
            data.append(row)
        #Generate table with datafor  PDF
        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), (0.8, 0.8, 0.8)),
            ('TEXTCOLOR', (0, 0), (-1, 0), (1, 1, 1)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), (0.95, 0.95, 0.95)),
        ]))
        elements.append(table)
    doc.build(elements)

# Function converting PDF to image
def convert_pdf_to_image(pdf_path, output_dir):
    images = convert_from_path(pdf_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    image_filenames = []
    for i, image in enumerate(images):
        image_filename = f"page_{i + 1}.png"
        image.save(os.path.join(output_dir, image_filename))
        image_filenames.append(image_filename)
    return image_filenames

#Return the images from directory
@app.route('/converted_images/<filename>')
def converted_images(filename):
    return send_from_directory(app.config['CONVERTED_FOLDER'], filename)

#Converting PDF to Excel file
def convert_pdf_to_xlsx(pdf_path, xlsx_path):
    # Open the PDF file
    with pdfplumber.open(pdf_path) as pdf:
        all_tables = []
        for page in pdf.pages:
            # Extract tables from the page
            tables = page.extract_tables()
            all_tables.extend(tables)
        # Using pandans dataframe for reading table
        df = pd.DataFrame([row for table in all_tables for row in table])
        # Save the dataframe to an excel file
        df.to_excel(xlsx_path, index=False)

#Converting PDF to Csv
def convert_pdf_to_csv(pdf_path, csv_path):
    tabula.convert_into(pdf_path, csv_path, output_format="csv")

@app.route('/converted_files/<filename>')
def converted_files(filename):
    return send_from_directory(app.config['CONVERTED_FOLDER'], filename)

#Converting PDF to text
def convert_pdf_to_text(pdf_path, output_path):
    text = extract_text(pdf_path)
    with open(output_path, "w", encoding="utf-8") as text_file:
        text_file.write(text)

# Rendering the converting function for converting supported formats
@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file:
        filename = file.filename
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        pdf_filename = f'{os.path.splitext(filename)[0]}.pdf'
        supported_extensions = {
            '.pptx', '.ppt', '.pps', '.ppsx', '.odp',
            '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.pbm', '.pgm', '.ppm', '.xbm', '.webp',
            '.doc', '.docx', '.docm', '.dot', '.dotm', '.dotx', '.odt', '.rtf', '.txt', '.wps', '.xml', '.xps',
            '.xlsx', '.xlsm', '.xltx', '.csv', '.pdf'
        }
        file_extension = os.path.splitext(filename)[-1].lower()
        if file_extension in supported_extensions:
            if file_extension == '.pdf':
                conversion_type = request.form['conversion_type']
                if conversion_type == 'pdf_to_image':
                    image_filenames = convert_pdf_to_image(file_path, app.config['CONVERTED_FOLDER'])
                    image_urls = [url_for('converted_images', filename=image_filename) for image_filename in
                                  image_filenames]
                    return render_template('index.html', image_urls=image_urls)
                elif conversion_type == 'pdf_to_xlsx':
                    xlsx_filename = f'{os.path.splitext(filename)[0]}.xlsx'
                    convert_pdf_to_xlsx(file_path, os.path.join(app.config['CONVERTED_FOLDER'], xlsx_filename))
                    xlsx_url = url_for('converted_files', filename=xlsx_filename)
                    return render_template('index.html', xlsx_url=xlsx_url)
                elif conversion_type == 'pdf_to_csv':
                    csv_filename = f'{os.path.splitext(filename)[0]}.csv'
                    convert_pdf_to_csv(file_path, os.path.join(app.config['CONVERTED_FOLDER'], csv_filename))
                    csv_url = url_for('converted_files', filename=csv_filename)
                    return render_template('index.html', csv_url=csv_url)
                elif conversion_type == 'pdf_to_text':
                    text_filename = f'{os.path.splitext(filename)[0]}.txt'
                    convert_pdf_to_text(file_path, os.path.join(app.config['CONVERTED_FOLDER'], text_filename))
                    text_url = url_for('converted_files', filename=text_filename)
                    return render_template('index.html', text_url=text_url)
                else:
                    return "Unsupported conversion type"
            else:
                if file_extension == '.pptx':
                    convert_ppt_to_pdf(file_path, os.path.join(app.config['CONVERTED_FOLDER'], pdf_filename))
                elif file_extension in (
                '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.pbm', '.pgm', '.ppm', '.xbm', '.webp'):
                    convert_image_to_pdf(file_path, os.path.join(app.config['CONVERTED_FOLDER'], pdf_filename))
                elif file_extension in (
                '.doc', '.docx', '.docm', '.dot', '.dotm', '.dotx', '.odt', '.rtf', '.txt', '.wps', '.xml', '.xps'):
                    convert_doc_to_pdf(file_path, os.path.join(app.config['CONVERTED_FOLDER'], pdf_filename))
                elif file_extension in ('.xlsx', '.xlsm', '.xltx'):
                    convert_excel_to_pdf(file_path, os.path.join(app.config['CONVERTED_FOLDER'], pdf_filename))
                elif file_extension == '.csv':
                    convert_csv_to_pdf(file_path, os.path.join(app.config['CONVERTED_FOLDER'], pdf_filename))
                pdf_url = url_for('uploaded_pdf', filename=pdf_filename)
                return render_template('index.html', pdf_url=pdf_url)
        else:
            return jsonify({"error": "Unsupported file format"})
    return jsonify({"error": "An error occurred"})

# Serving the folders for file conversion
@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/converted_pdfs/<filename>')
def uploaded_pdf(filename):
    return send_from_directory(app.config['CONVERTED_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
