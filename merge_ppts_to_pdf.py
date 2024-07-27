import os
import subprocess
from PyPDF2 import PdfMerger

def convert_ppt_to_pdf(ppt_path, pdf_path):
    subprocess.run([
        'soffice',
        '--headless',
        '--convert-to',
        'pdf',
        '--outdir',
        os.path.dirname(pdf_path),
        ppt_path
    ], check=True)

def merge_pdfs(pdf_paths, output_path):
    merger = PdfMerger()
    for pdf in pdf_paths:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()

def main(ppt_folder, output_pdf):
    ppt_files = [f for f in os.listdir(ppt_folder) if f.endswith('.ppt') or f.endswith('.pptx')]
    pdf_files = []

    for ppt in ppt_files:
        ppt_path = os.path.join(ppt_folder, ppt)
        pdf_path = os.path.join(ppt_folder, ppt.replace('.pptx', '.pdf').replace('.ppt', '.pdf'))
        convert_ppt_to_pdf(ppt_path, pdf_path)
        pdf_files.append(pdf_path)

    merge_pdfs(pdf_files, output_pdf)

    # Clean up individual PDF files if needed
    for pdf in pdf_files:
        os.remove(pdf)

if __name__ == "__main__":
    
    out_file_name = 'merged.pdf'
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    ppt_folder = os.path.join(script_dir, 'to_merge')  
    output_pdf = os.path.join(script_dir, 'output', out_file_name)   
    main(ppt_folder, output_pdf)