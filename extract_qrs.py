import sys
import re
import os
import traceback
from pypdf import PdfReader

def extract_qrs(pdf_path, output_path, regex_pattern):
    try:
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
            
        pattern = re.compile(regex_pattern)
        reader = PdfReader(pdf_path)
        unique_qrs = set()
        
        total_pages = len(reader.pages)
        progress_path = output_path + ".prog"
        for idx, page in enumerate(reader.pages):
            # Write current progress file (non-blocking IPC)
            try:
                with open(progress_path, "w", encoding="utf-8") as pf:
                    pf.write(f"{idx+1}/{total_pages}")
            except:
                pass
                
            text = page.extract_text()
            if text:
                matches = pattern.findall(text)
                for match in matches:
                    unique_qrs.add(match.strip())
                    
        # Delete progress file when done
        if os.path.exists(progress_path):
            try:
                os.remove(progress_path)
            except:
                pass
                    
        # Write unique extracted values to temporary file
        with open(output_path, "w", encoding="utf-8") as f:
            for qr in sorted(unique_qrs):
                f.write(qr + "\n")
                
        print(f"SUCCESS: Extracted {len(unique_qrs)} unique QR values.")
        sys.exit(0)
    except Exception as e:
        # Write diagnostics traceback to parser_log.txt in same directory as temp file
        log_dir = os.path.dirname(output_path)
        log_path = os.path.join(log_dir, "parser_log.txt")
        try:
            with open(log_path, "w", encoding="utf-8") as lf:
                traceback.print_exc(file=lf)
        except:
            pass
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: extract_qrs.py <pdf_path> <output_path> <regex_pattern>")
        sys.exit(1)
    extract_qrs(sys.argv[1], sys.argv[2], sys.argv[3])
