import os
import sys
import re
import fitz  # PyMuPDF
import easyocr
import numpy as np
import cv2
from datetime import datetime

reader = easyocr.Reader(['en'])

DATE_PATTERN = r"\b\d{1,2}[-/\s]\d{1,2}[-/\s]\d{2,4}\b"
TIME_PATTERN = r"\b\d{1,2}:\d{2}(:\d{2})?\b"

CN_PATTERN = r"(CN[-_\s]?[A-Z0-9]+[-_\s]?[0-9]+)"
PO_STANDARD_PATTERN = r"(P[O|0][0-9]{7,8})"
PO_CUT_PATTERN = r"([O|0][0-9]{7})"

def clean_and_extract_id(page_text, filename):
    top_rows = page_text[:12]
    full_text = " ".join(top_rows).upper()
    
    print(f"   [OCR Text inside {filename}]: {full_text if full_text.strip() else '[No Text]'}")
    
    full_text = re.sub(DATE_PATTERN, "", full_text)
    full_text = re.sub(TIME_PATTERN, "", full_text)
    
    cn_match = re.search(CN_PATTERN, full_text)
    if cn_match:
        return cn_match.group(1).replace(" ", "")
        
    po_match = re.search(PO_STANDARD_PATTERN, full_text)
    if po_match:
        clean_po = po_match.group(1).replace(" ", "")
        if clean_po.startswith("P0"):
            clean_po = "PO" + clean_po[2:]
        return clean_po
        
    po_cut_match = re.search(PO_CUT_PATTERN, full_text)
    if po_cut_match:
        captured_id = po_cut_match.group(1).replace(" ", "")
        return f"PO{captured_id[1:]}"
        
    return None

def extract_id_from_pdf(pdf_path, filename):
    try:
        doc = fitz.open(pdf_path)
        if len(doc) == 0:
            return None
            
        page = doc[0]
        zoom = 3  
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.h, pix.w, pix.n))
        img_rgb = cv2.cvtColor(img_data, cv2.COLOR_BGR2RGB)
        
        page_text = reader.readtext(img_rgb, detail=0)
        doc.close()
        
        return clean_and_extract_id(page_text, filename)
    except Exception as e:
        print(f"Error reading {filename}: {e}")
    return None

def main(target_folder):
    if not os.path.exists(target_folder):
        print(f"Error: Folder not found at {target_folder}")
        return

    today_str = datetime.now().strftime("%d %B")
    output_folder = os.path.join(target_folder, today_str)
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    pdf_files = [f for f in os.listdir(target_folder) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"No PDF files found in {target_folder}")
        return

    print(f"Target Folder: {target_folder}")
    print(f"Output Directory: {output_folder}")
    print(f"Processing {len(pdf_files)} single PDFs...\n")

    for filename in pdf_files:
        if filename == today_str:
            continue
            
        file_path = os.path.join(target_folder, filename)
        print(f"Scanning: {filename}...")
        
        doc_id = extract_id_from_pdf(file_path, filename)
        
        if doc_id:
            new_filename = f"{doc_id}.pdf"
            new_file_path = os.path.join(output_folder, new_filename)
            
            counter = 1
            while os.path.exists(new_file_path):
                new_filename = f"{doc_id}_{counter}.pdf"
                new_file_path = os.path.join(output_folder, new_filename)
                counter += 1
                
            os.rename(file_path, new_file_path)
            print(f"--> SUCCESS: Moved & Renamed to -> {today_str}/{new_filename}\n")
        else:
            print(f"--> FAILED: No valid ID found in this file.\n")
            
    print("All single documents processed successfully.")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        folder_to_process = sys.argv[1]
    else:
        folder_to_process = "./"
        
    main(folder_to_process)
    os.system("pause")
