import csv
from docx import Document
import os
import subprocess
import time

def doc2pdf_linux(doc, pdf):
    """
    Convert a doc/docx document to PDF format (Linux only, requires LibreOffice).
    :param doc: Path to the document
    """
    cmd = ['libreoffice', '--convert-to', 'pdf', doc, '--outdir', os.path.dirname(pdf)]
    p = subprocess.Popen(cmd, stderr=subprocess.PIPE, stdout=subprocess.PIPE, env={"HOME": "/tmp"})
    p.wait(timeout=10)
    stdout, stderr = p.communicate()
    if stderr:
        raise subprocess.SubprocessError(stderr)


# Load your list of names from a CSV file
with open('Participants_List.csv', 'r') as f:
    reader = csv.reader(f)
    names = []
    names_underscore = []
    for row in reader:
        name = row[0]
        names.append(name)
        names_underscore.append(name.replace(' ', '_'))

for name1, name2 in zip(names, names_underscore):
    # Load your certificate template
    doc = Document('Template.docx')

    # Replace placeholders with actual values
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if 'XYZ' in run.text:
                run.text = run.text.replace('XYZ', name1)
            if "ABC" in run.text:
                run.text = run.text.replace("ABC", "Chamel Nadir Bouacha")  # Replace ABC with the Ambassador Name
            if "CSC" in run.text:
                run.text = run.text.replace("CSC", "Ambassador Challenge Dive Into The World Of Cloud With Azure")  # Replace CSC with the Event Name
            '''if "DATE" in run.text:
                run.text = run.text.replace("DATE", "8th March 2024")  # Replace DATE with the Event Date'''

    # Create the necessary directories if they don't exist
    os.makedirs('Output/DOCS/', exist_ok=True)
    os.makedirs('Output/PDF/', exist_ok=True)

    # Save the certificate as .docx
    docs = f'Output/DOCS/{name2}_certificate.docx'
    pdf = f'Output/PDF/{name2}_certificate.pdf'
    doc.save(docs)

    time.sleep(0.1)  # To prevent the program from crashing due to the file not being saved yet

    # Convert the .docx file to .pdf
    doc2pdf_linux(docs, pdf)

print("Certificates generated and converted to PDF successfully!")
