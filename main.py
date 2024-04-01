import csv
from docx import Document
import os
from docx2pdf import convert

# Load your list of names from a CSV file
with open('Participants_List.csv', 'r') as f:
    reader = csv.reader(f)
    names = [row[0] for row in reader]  # adjust this based on the structure of your CSV file

for name in names:
    # Load your certificate template
    doc = Document('Template.docx')

    # Replace placeholder with the actual name
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if 'XYZ' in run.text:
                run.text = run.text.replace('XYZ', name)
            if "ABC" in run.text:
                run.text = run.text.replace("ABC", "Chamel Nadir Bouacha") # Replace ABC With The Ambassador Name
            if "CSC" in run.text:
                run.text = run.text.replace("CSC", "Ambassador Challenge Dive Into The World Of Cloud With Azure") # Replace CSC With The Event Name

    # Create the necessary directories if they don't exist
    os.makedirs('Output/DOCS/', exist_ok=True)
    os.makedirs('Output/PDF/', exist_ok=True)

    # Save the certificate
    docs = f'Output/DOCS/{name}_certificate.docx'
    pdf = f'Output/PDF/{name}_certificate.pdf'
    doc.save(docs)

    # Convert the .docx file to .pdf
    convert(docs, pdf)
