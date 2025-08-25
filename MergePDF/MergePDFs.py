import os
from docx2pdf import convert
from PyPDF2 import PdfMerger

input_folder = r"."
output_file = r".\combined.pdf"

temp_pdf_folder = os.path.join(input_folder, "temp_pdfs")
os.makedirs(temp_pdf_folder, exist_ok=True)

for file in os.listdir(input_folder):
    if file.lower().endswith(".docx"):
        word_path = os.path.join(input_folder, file)
        pdf_path = os.path.join(temp_pdf_folder, f"{os.path.splitext(file)[0]}.pdf")
        convert(word_path, pdf_path)

merger = PdfMerger()

for file in sorted(os.listdir(temp_pdf_folder)):
    if file.lower().endswith(".pdf"):
        merger.append(os.path.join(temp_pdf_folder, file))

merger.write(output_file)
merger.close()

print(f"Merged PDF saved to {output_file}")