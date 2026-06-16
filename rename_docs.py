import os
import sys
import re
import fitz  # PyMuPDF
import easyocr
import numpy as np
import cv2

reader = easyocr.Reader(['en'])

CN_PATTERN = r"(CN-[A-Za-z0-9]+-[0-9]+)"
PO_PATTERN = r"(PO[0-9]+)"

def extract_id_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        if len(doc) == 0:
            return None
            
        page = doc[0]
        zoom = 2  
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.h, pix.w, pix.n))
        img_rgb = cv2.cvtColor(img_data, cv2.COLOR_BGR2RGB)
        
        results = reader.readtext(img_rgb, detail=0)
        full_text = " ".join(results).upper()
        
        cn_match = re.search(CN_PATTERN, full_text)
        if cn_match:
            return cn_match.group(1).replace(" ", "")
            
        po_match = re.search(PO_PATTERN, full_text)
        if po_match:
            return po_match.group(1).replace(" ", "")
            
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
        
    return None

def process_and_rename_pdfs(target_folder):
    if not os.path.exists(target_folder):
        print(f"Error: Folder not found at {target_folder}")
        return

    pdf_files = [f for f in os.listdir(target_folder) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"No PDF files found in {target_folder}")
        return

    print(f"Target Folder: {target_folder}")
    print(f"Found {len(pdf_files)} PDFs to process...")

    for filename in pdf_files:
        file_path = os.path.join(target_folder, filename)
        print(f"\nScanning PDF: {filename}...")
        
        doc_id = extract_id_from_pdf(file_path)
        
        if doc_id:
            new_filename = f"{doc_id}.pdf"
            new_file_path = os.path.join(target_folder, new_filename)
            
            if not os.path.exists(new_file_path):
                os.rename(file_path, new_file_path)
                print(f"--> SUCCESS: Renamed to -> {new_filename}")
            else:
                print(f"--> SKIPPED: {new_filename} already exists.")
        else:
            print(f"--> FAILED: No valid ID pattern found in this document.")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        folder_to_process = sys.argv[1]
    else:
        folder_to_process = "./"
        
    process_and_rename_pdfs(folder_to_process)
    print("\nProcessing complete.")
    os.system("pause")