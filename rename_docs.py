import os
import sys
import re
import time
import fitz  # PyMuPDF
import easyocr
import numpy as np
import cv2
from datetime import datetime

# Initialize EasyOCR Reader once
reader = easyocr.Reader(['en'])

# Regex Patterns
DATE_PATTERN = r"\b\d{1,2}[-/\s]\d{1,2}[-/\s]\d{2,4}\b"
TIME_PATTERN = r"\b\d{1,2}:\d{2}(:\d{2})?\b"

CN_PATTERN = r"(CN[-_\s]?[A-Z0-9]+[-_\s]?[0-9]+)"
PO_STANDARD_PATTERN = r"(P[O|0][0-9]{7,8})"
PO_CUT_PATTERN = r"([O|0][0-9]{7})"

def clean_and_extract_id(text_string):
    """ Centralized extraction logic for a single string of text """
    if not text_string:
        return None
        
    upper_text = text_string.upper()
    # Remove dates and times to prevent false matching
    upper_text = re.sub(DATE_PATTERN, "", upper_text)
    upper_text = re.sub(TIME_PATTERN, "", upper_text)
    
    # 1. Match Standard CN Pattern
    cn_match = re.search(CN_PATTERN, upper_text)
    if cn_match:
        return cn_match.group(1).replace(" ", "")
        
    # 2. Match Standard PO Pattern
    po_match = re.search(PO_STANDARD_PATTERN, upper_text)
    if po_match:
        clean_po = po_match.group(1).replace(" ", "")
        if clean_po.startswith("P0"):
            clean_po = "PO" + clean_po[2:]
        return clean_po
        
    # 3. Match Cut-off PO Pattern
    po_cut_match = re.search(PO_CUT_PATTERN, upper_text)
    if po_cut_match:
        captured_id = po_cut_match.group(1).replace(" ", "")
        return f"PO{captured_id[1:]}"
        
    return None

def extract_id_from_pdf(pdf_path, filename):
    try:
        doc = fitz.open(pdf_path)
        if len(doc) == 0:
            return None
            
        # 🏎️ STRATEGY 1: Digital Text Fast-Scan (Takes ~0.005 seconds)
        # If the PDF contains embedded text layer, grab it instantly without OCR
        raw_text = ""
        for page in doc:
            raw_text += page.get_text()
            
        digital_id = clean_and_extract_id(raw_text)
        if digital_id:
            doc.close()
            print("   [Fast-Track]: ID found instantly via Digital Text Layer!")
            return digital_id
            
        # 📷 STRATEGY 2: Single-Shot Smart Image OCR (If digital text layer fails)
        page = doc[0]
        zoom = 3  
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.h, pix.w, pix.n))
        doc.close()
        
        # We divide into stable 2cm blocks vertically, but keep FULL WIDTH (No column splitting)
        # 15 rows total for A4 height. We only scan the top 4 rows (~8cm total space)
        chunk_height = int(pix.h / 15) 
        
        for row in range(4): 
            start_y = row * chunk_height
            end_y = (row + 1) * chunk_height
            
            # Crop Full Width to keep the layout unbroken
            cropped_cell = img_data[start_y:end_y, 0:pix.w]
            img_rgb = cv2.cvtColor(cropped_cell, cv2.COLOR_BGR2RGB)
            
            # Single OCR pass per block
            page_text = reader.readtext(img_rgb, detail=0)
            
            if page_text:
                combined_word = " ".join(page_text).strip()
                if combined_word:
                    print(f"   [Scan Block {row+1}]: {combined_word[:40]}")
                    detected_id = clean_and_extract_id(combined_word)
                    if detected_id:
                        return detected_id
                            
    except Exception as e:
        print(f"   Error reading {filename}: {e}")
    return None

def main(target_folder):
    if not os.path.exists(target_folder):
        print(f"Error: Folder not found at {target_folder}")
        return

    start_all_time = time.time()

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
    print(f"Processing {len(pdf_files)} PDFs with Digital Hybrid Engine...\n")

    table_data = []
    skipped_count = 0

    for filename in pdf_files:
        if filename == today_str:
            continue
            
        start_file_time = time.time()
        name_upper = filename.upper()
        
        if name_upper.startswith("PO") or name_upper.startswith("CN"):
            print(f"Skipping: {filename} (Already renamed)")
            src_path = os.path.join(target_folder, filename)
            dst_path = os.path.join(output_folder, filename)
            if not os.path.exists(dst_path):
                os.rename(src_path, dst_path)
            
            end_file_time = time.time()
            table_data.append({
                "original": filename,
                "renamed": filename,
                "time": f"{end_file_time - start_file_time:.2f}s",
                "status": "Skipped"
            })
            skipped_count += 1
            continue
            
        file_path = os.path.join(target_folder, filename)
        print(f"Scanning: {filename}...")
        
        doc_id = extract_id_from_pdf(file_path, filename)
        
        end_file_time = time.time()
        file_elapsed = end_file_time - start_file_time
        
        if doc_id:
            new_filename = f"{doc_id}.pdf"
            new_file_path = os.path.join(output_folder, new_filename)
            
            counter = 1
            while os.path.exists(new_file_path):
                new_filename = f"{doc_id}_{counter}.pdf"
                new_file_path = os.path.join(output_folder, new_filename)
                counter += 1
                
            os.rename(file_path, new_file_path)
            print(f"--> SUCCESS: Renamed to -> {today_str}/{new_filename} [Time: {file_elapsed:.2f}s]\n")
            
            table_data.append({
                "original": filename,
                "renamed": new_filename,
                "time": f"{file_elapsed:.2f}s",
                "status": "Success"
            })
        else:
            print(f"--> FAILED: No ID found. [Time: {file_elapsed:.2f}s]\n")
            table_data.append({
                "original": filename,
                "renamed": "FAILED (No ID)",
                "time": f"{file_elapsed:.2f}s",
                "status": "Failed"
            })
            
    end_all_time = time.time()
    total_elapsed_time = end_all_time - start_all_time
    
    # Print Final Summary Table
    print("\n" + "="*85)
    print(f" {'ORIGINAL FILE NAME':<25} | {'RENAMED FILE NAME':<30} | {'TIME TAKEN':<12} | {'STATUS':<10}")
    print("="*85)
    for row in table_data:
        orig = row['original'][:23] + ".." if len(row['original']) > 25 else row['original']
        renamed = row['renamed'][:28] + ".." if len(row['renamed']) > 30 else row['renamed']
        print(f" {orig:<25} | {renamed:<30} | {row['time']:<12} | {row['status']:<10}")
    print("="*85)
    print(f" Total Files Processed : {len(table_data) - skipped_count}")
    print(f" Total Files Skipped   : {skipped_count}")
    print(f" Total Time Taken      : {total_elapsed_time:.2f} seconds")
    print("="*85 + "\n")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        folder_to_process = sys.argv[1]
    else:
        folder_to_process = "./"
        
    main(folder_to_process)
    os.system("pause")
