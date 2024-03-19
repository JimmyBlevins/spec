from pathlib import Path
import docx
from docx2pdf import convert
from PyPDF2 import PdfMerger

def spec_update(data):

    file_path = Path(data['Folder'])

    # Iterate over each .docx file in directory
    for file in file_path.glob('*.docx'):
        doc = docx.Document(file)
        doc.core_properties.category = data["ClientName"]
        doc.core_properties.keywords = data["ProjectName"]
        doc.core_properties.content_status = data["ProjectStatus"]
        doc.core_properties.comments = data["ProjectNo"]
        doc.save(file)

    #convert files to pdf
    if data["CreatePDFs"]:
        convert(file_path)

        if data["MergePDFs"]:
            merge(file_path,data["ProjectStatus"])
    
    return

def merge(file_path,status):

    merger = PdfMerger()

    # Iterate over each PDF file in the directory
    for file in file_path.glob('*.pdf'):
        with open(file, 'rb') as pdf_file:
            merger.append(pdf_file)

    # Specify the output file path
    output_path = file_path / f'merged_specs ({status}).pdf'
        
    # Merge the PDF files
    with open(output_path, 'wb') as output_file:
        merger.write(output_file)
    
    # Close the merger
    merger.close()

    return 