import os
import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from io import BytesIO
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.utils import simpleSplit

# Register a Unicode CID font
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))

def get_pdf_files(directory):
    pdf_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    return pdf_files

def create_toc(pdf_files, page_numbers):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    c.setFont("Helvetica-Bold", 18)
    c.drawString(1*inch, height - 1*inch, "Family Relocation Documents - Table of Contents")

    y_position = height - 2*inch
    c.setFont("HeiseiMin-W3", 12)

    for i, (pdf_file, page_num) in enumerate(zip(pdf_files, page_numbers), start=1):
        relative_path = os.path.relpath(pdf_file)

        if y_position < 1*inch:  # Start a new page if we're near the bottom
            c.showPage()
            y_position = height - 1*inch

        text = f"{i}. {relative_path}"
        for line in simpleSplit(text, 'HeiseiMin-W3', 12, 6*inch):
            c.drawString(1*inch, y_position, line)
            y_position -= 0.2*inch
        c.drawRightString(7.5*inch, y_position, str(page_num))
        y_position -= 0.3*inch

    c.save()
    buffer.seek(0)
    return PyPDF2.PdfReader(buffer)

def create_title_page(title):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    c.setFont("HeiseiMin-W3", 18)
    # Split long titles into multiple lines
    lines = simpleSplit(title, 'HeiseiMin-W3', 18, width - 2*inch)
    y = height/2 + (len(lines) * 24 / 2)  # Start higher for multi-line titles
    for line in lines:
        c.drawCentredString(width/2, y, line)
        y -= 24  # Move down for next line

    c.save()
    buffer.seek(0)
    return PyPDF2.PdfReader(buffer)

def merge_pdfs(pdf_files, output_path):
    merger = PyPDF2.PdfMerger()

    # First, calculate page numbers for ToC
    page_numbers = [1]  # ToC starts at page 1
    for pdf_file in pdf_files:
        try:
            with open(pdf_file, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                page_numbers.append(page_numbers[-1] + len(reader.pages) + 1)  # +1 for title page
        except Exception:
            page_numbers.append(page_numbers[-1] + 1)  # Error page

    # Add table of contents
    toc = create_toc(pdf_files, page_numbers[1:])
    merger.append(toc)

    # Add each PDF
    for pdf_file in pdf_files:
        try:
            # Add a title page
            title = os.path.relpath(pdf_file)
            title_page = create_title_page(title)
            merger.append(title_page)

            # Add the actual PDF
            merger.append(pdf_file)

            # Add bookmark
            merger.add_outline_item(title, page_numbers[pdf_files.index(pdf_file)], parent=None)

        except Exception as e:
            print(f"Error processing {pdf_file}: {str(e)}")
            # Create a page with error message
            buffer = BytesIO()
            c = canvas.Canvas(buffer, pagesize=letter)
            c.setFont("HeiseiMin-W3", 14)
            c.drawString(1*inch, 9*inch, f"Error processing: {os.path.relpath(pdf_file)}")
            c.setFont("HeiseiMin-W3", 12)
            c.drawString(1*inch, 8*inch, f"Error: {str(e)}")
            c.save()
            buffer.seek(0)
            merger.append(PyPDF2.PdfReader(buffer))

    # Write the merged PDF to disk
    merger.write(output_path)
    merger.close()

def main():
    current_dir = os.getcwd()
    pdf_files = get_pdf_files(current_dir)
    output_path = os.path.join(current_dir, 'MergedFamilyRelocationDocuments.pdf')
    merge_pdfs(pdf_files, output_path)
    print(f"Merged PDF created successfully: {output_path}")

if __name__ == "__main__":
    main()
