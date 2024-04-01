import csv
from docx import Document
import os

# Load your list of names from a CSV file
with open('Participants_List.csv', 'r') as f:
    reader = csv.reader(f)
    names = [row[0] for row in reader]  # adjust this based on the structure of your CSV file

# Load your certificate template
doc = Document('Template.odt')

for name in names:
    # Make a copy of the template
    certificate = doc.copy()

    # Replace placeholder with the actual name
    for paragraph in certificate.paragraphs:
        if 'XYZ' in paragraph.text:
            paragraph.text = paragraph.text.replace('XYZ', name)
        if "ABC" in paragraph.text:
            paragraph.text = paragraph.text.replace("ABC", "Chamel Nadir Bouacha") # Replace ABC With The Ambassador Name
        if "CSC" in paragraph.text:
            paragraph.text = paragraph.text.replace("CSC", "Ambassador Challenge Dive Into The World Of Cloud With Azure") # Replace CSC With The Event Name
    # Save the certificate
    odt = f'Output/ODT/{name}_certificate.odt'
    pdf = f'Output/PDF/{name}_certificate.pdf'
    certificate.save(f'Output/ODT/{name}_certificate.odt')
        # Convert the .odt file to .pdf
    os.system(f'unoconv -f pdf -o {pdf} {odt}')
