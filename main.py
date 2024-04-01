import csv
from docx import Document

# Load your list of names from a CSV file
with open('path_to_your_csv_file.csv', 'r') as f:
    reader = csv.reader(f)
    names = [row[0] for row in reader]  # adjust this based on the structure of your CSV file

# Load your certificate template
doc = Document('path_to_your_certificate_template.odt')

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
    certificate.save(f'certificates/{name}_certificate.odt')
