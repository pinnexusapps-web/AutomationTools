import os
import sys
import re
import fitz  # PyMuPDF
import easyocr
import numpy as np
import cv2

reader = easyocr.Reader(['en'])

# 🛠️ Date සහ Time කොටස් හඳුනාගෙන ඉවත් කිරීමට සාදන ලද Regex රටාවන්:
# 16/06/2026 හෝ 16-06-2026 වැනි Dates
DATE_PATTERN = r"\b\d{1,2}[-/\s]\d{1,2}[-/\s]\d{2,4}\b"
# 12:22 හෝ 14:35:00 වැනි Times
TIME_PATTERN = r"\b\d{1,2}:\d{2}(:\d{2})?\b"

# Document ID Regex Patterns
CN_PATTERN = r"(CN[-_\s]?[A-Z0-9]+[-_\s]?[0-9]+)"
PO_STANDARD_PATTERN = r"(PO[0-9]{7,8})"
PO_CUT_PATTERN = r"([O|0][0-9]{7})"

def clean_and_extract_id(results):
    # 1. OCR වචන සියල්ල එකතු කර තනි පේළියක් සාදා ගැනීම
    full_text = " ".join(results).upper()
    
    # 2. Date සහ Time කොටස් තිබේ නම් ඒවා සම්පූර්ණයෙන්ම ඉවත් කිරීම (Cleaning)
    full_text = re.sub(DATE_PATTERN, "", full_text)
    full_text = re.sub(TIME_PATTERN, "", full_text)
    
    # 3. Standard CN Pattern එක සෙවීම
    cn_match = re.search(CN_PATTERN, full_text)
    if cn_match:
        return cn_match.group(1).replace(" ", "")
        
    # 4. Standard PO Pattern එක සෙවීම (උදා: PO4198698)
    po_match = re.search(PO_STANDARD_PATTERN, full_text)
    if po_match:
        clean_po = po_match.group(1).replace(" ", "")
        if clean_po.startswith("P0"):
            clean_po = "PO" + clean_po[2:]
        return clean_po
        
    # 5. Border කැපී ඉතිරි වූ 'O' + ඉලක්කම් 7 රටාව සෙවීම (උදා: O4200077)
    po_cut_match = re.search(PO_CUT_PATTERN, full_text)
    if po_cut_match:
        captured_id = po_cut_match.group(1).replace(" ", "")
        return f"PO{captured_id[1:]}"
        
    return None

def extract_id_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        if len(doc) == 0:
            return None
            
        page = doc[0]
        zoom = 3  # High resolution for precise scanning
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.h, pix.w, pix.n))
        img_rgb = cv2.cvtColor(img_data, cv2.COLOR_BGR2RGB)
        
        results = reader.readtext(img_rgb, detail=0)
        return clean_and_extract_id(results)
            
    except Exception as e:
        print(f"Error inside OCR processing: {e}")
        
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
    print(f"Processing {len(pdf_files)} PDFs using Date/Time Filtering OCR...")

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
