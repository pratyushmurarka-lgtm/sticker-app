import re
import os
from pypdf import PdfReader, PdfWriter

def scan_pdf_for_qrs(pdf_path, regex_pattern):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF file not found at path: {pdf_path}")
        
    pattern = re.compile(regex_pattern)
    reader = PdfReader(pdf_path)
    unique_qrs = set()
    
    for page in reader.pages:
        text = page.extract_text()
        if text:
            matches = pattern.findall(text)
            for match in matches:
                unique_qrs.add(match.strip())
                
    return sorted(list(unique_qrs))

def split_pdf_pages(pdf_path, output_path, from_page, to_page):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF file not found at path: {pdf_path}")
        
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    total_pages = len(reader.pages)
    
    # 1-based to 0-based conversion
    start = max(0, from_page - 1)
    end = min(total_pages, to_page)
    
    if start >= total_pages or start > end:
        raise ValueError(f"Invalid page range: {from_page} to {to_page}. Total pages: {total_pages}")
        
    for idx in range(start, end):
        writer.add_page(reader.pages[idx])
        
    with open(output_path, "wb") as out_f:
        writer.write(out_f)
        
    return output_path
