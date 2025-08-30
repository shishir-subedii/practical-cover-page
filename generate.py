import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
from PyPDF2 import PdfMerger
import os
import shutil

# ==== CONFIG ====
TEMPLATE_FILE = "template.docx"
SPREADSHEET = "subjects.xlsx"
OUTPUT_DIR = "output"
INDEX_FILE = "index.pdf"

# === Reset output folder ===
if os.path.exists(OUTPUT_DIR):
    shutil.rmtree(OUTPUT_DIR)  # delete old output + contents
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Ask student name
student_name = input("Enter your name: ")

# Load subjects from spreadsheet
df = pd.read_excel(SPREADSHEET)

generated_pdfs = []

# Loop through each row (subject, teacher)
for _, row in df.iterrows():
    subject = row["Subject_Name"]
    teacher = row["Teacher_Name"]

    # Load template
    doc = DocxTemplate(TEMPLATE_FILE)

    # Context to replace placeholders
    context = {
        "Subject_Name": subject,
        "Teacher_Name": teacher,
        "Student_Name": student_name
    }

    # File paths
    filename_docx = os.path.join(OUTPUT_DIR, f"{subject}.docx")
    filename_pdf = os.path.join(OUTPUT_DIR, f"{subject}.pdf")

    # Render and save DOCX
    doc.render(context)
    doc.save(filename_docx)

    # Convert to PDF
    convert(filename_docx, filename_pdf)
    print(f"Generated: {filename_pdf}")

    generated_pdfs.append(filename_pdf)

    # Delete temporary DOCX
    os.remove(filename_docx)

# Copy index.pdf into /output
index_copy = os.path.join(OUTPUT_DIR, "index.pdf")
shutil.copy(INDEX_FILE, index_copy)

# === Merge PDFs ===
merger = PdfMerger()

# Add index.pdf first
merger.append(index_copy)

# Add all subject PDFs (order from Excel)
for pdf in generated_pdfs:
    merger.append(pdf)

# Save combined file
combined_pdf = os.path.join(OUTPUT_DIR, "combined_pages.pdf")
merger.write(combined_pdf)
merger.close()

print("Done! Combined file created at:", combined_pdf)