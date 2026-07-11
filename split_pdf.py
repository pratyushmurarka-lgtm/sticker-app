import sys
import os
from pypdf import PdfReader, PdfWriter

def split_pdf(pdf_path, output_path, from_page, to_page):
    try:
        if not os.path.exists(pdf_path):
            print(f"ERROR: Input file not found: {pdf_path}")
            return
            
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        total_pages = len(reader.pages)
        
        # Adjust 1-based indexing to 0-based indexing
        start = max(0, from_page - 1)
        end = min(total_pages, to_page)
        
        if start >= total_pages or start > end:
            print("ERROR: Invalid page range")
            return
            
        for idx in range(start, end):
            writer.add_page(reader.pages[idx])
            
        with open(output_path, "wb") as out_f:
            writer.write(out_f)
            
        print("SUCCESS")
    except Exception as e:
        print(f"ERROR: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("ERROR: Missing arguments")
        sys.exit(1)
        
    pdf_path = sys.argv[1]
    output_path = sys.argv[2]
    from_page = int(sys.argv[3])
    to_page = int(sys.argv[4])
    
    split_pdf(pdf_path, output_path, from_page, to_page)
