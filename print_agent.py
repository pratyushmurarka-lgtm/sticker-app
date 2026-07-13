import os
import sys
import json
import base64
import tempfile
from http.server import HTTPServer, BaseHTTPRequestHandler
import win32print
import win32ui
import win32con
import win32api
from PIL import Image, ImageWin
import segno

class CORSHTTPRequestHandler(BaseHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(204)
        self.end_headers()

    def do_GET(self):
        if self.path == "/printers":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            
            # Get list of installed printers
            printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            default_printer = win32print.GetDefaultPrinter()
            
            self.wfile.write(json.dumps({
                "printers": printers,
                "default": default_printer
            }).encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == "/print":
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            
            try:
                data = json.loads(post_data.decode('utf-8'))
                mode = data.get("mode")
                printer_name = data.get("printer_name") or win32print.GetDefaultPrinter()
                
                if mode == "pdf":
                    pdf_base64 = data.get("pdf_data")
                    if not pdf_base64:
                        raise ValueError("Missing pdf_data for pdf mode")
                    success = self.print_pdf(pdf_base64, printer_name)
                    
                elif mode == "sequential":
                    qrs = data.get("qrs", [])
                    label_size_mm = float(data.get("label_size", 25))
                    gap_mm = float(data.get("gap", 3))
                    top_margin_mm = float(data.get("top_margin", 5))
                    success = self.print_sequential(qrs, printer_name, label_size_mm, gap_mm, top_margin_mm)
                else:
                    raise ValueError(f"Unknown mode: {mode}")
                    
                if success:
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"success": True, "message": "Printed successfully"}).encode('utf-8'))
                else:
                    raise RuntimeError("Printing failed")
                    
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"success": False, "message": str(e)}).encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

    def print_pdf(self, pdf_base64, printer_name):
        # Decode and write to temporary file
        pdf_bytes = base64.b64decode(pdf_base64)
        fd, temp_pdf_path = tempfile.mkstemp(suffix=".pdf")
        try:
            with os.fdopen(fd, "wb") as f:
                f.write(pdf_bytes)
                
            # Set default printer dynamically (equivalent to SetDefaultPrinter in VB6)
            original_printer = win32print.GetDefaultPrinter()
            if original_printer != printer_name:
                win32print.SetDefaultPrinter(printer_name)
                
            # Send print command using ShellExecute
            win32api.ShellExecute(0, "print", temp_pdf_path, None, ".", 0)
            
            # Revert default printer if changed
            if original_printer != printer_name:
                # Wait 2 seconds for shell print to spool before reverting default printer
                win32api.Sleep(2000)
                win32print.SetDefaultPrinter(original_printer)
                
            return True
        finally:
            # We delay deletion of the temp file to allow spooler to read it
            def cleanup():
                try:
                    win32api.Sleep(5000)
                    os.remove(temp_pdf_path)
                except:
                    pass
            import threading
            threading.Thread(target=cleanup).start()

    def print_sequential(self, qrs, printer_name, label_size_mm, gap_mm, top_margin_mm):
        # 1. Setup printer Device Context
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        
        # Get printer resolutions (DPI)
        dpi_x = hdc.GetDeviceCaps(win32con.LOGPIXELSX)
        dpi_y = hdc.GetDeviceCaps(win32con.LOGPIXELSY)
        
        # Conversion factors: mm to dots
        mm_to_dot_x = dpi_x / 25.4
        mm_to_dot_y = dpi_y / 25.4
        
        label_size_dots = int(label_size_mm * mm_to_dot_x)
        gap_dots = int(gap_mm * mm_to_dot_x)
        top_margin_dots = int(top_margin_mm * mm_to_dot_y)
        left_margin_dots = int(3.0 * mm_to_dot_x) # 3mm left margin
        
        # Prepare layout variables
        total_qrs = len(qrs)
        labels_per_page = 4
        total_pages = (total_qrs + labels_per_page - 1) // labels_per_page
        
        hdc.StartDoc("Sequential QR Labels")
        
        for p in range(total_pages):
            hdc.StartPage()
            
            # Select Font (Calibri, size 8 equivalent)
            font_height = int(3.0 * mm_to_dot_y) # ~8pt
            font = win32ui.CreateFont({
                "name": "Calibri",
                "height": font_height,
                "weight": win32con.FW_NORMAL,
            })
            hdc.SelectObject(font)
            
            for i in range(labels_per_page):
                idx = p * labels_per_page + i
                if idx >= total_qrs:
                    break
                    
                qr_value = qrs[idx]
                
                # Format text: CodeStr-SerialNo
                parts = qr_value.split("|")
                if len(parts) >= 7:
                    label_text = f"{parts[2].strip()}-{parts[6].strip()}"
                else:
                    label_text = qr_value
                    
                # 2. Draw QR code
                # Generate QR code PIL image using segno
                qr = segno.make(qr_value, error='M')
                # Scale up QR code to label size for crisp printing
                qr_img = qr.to_pil(scale=10, border=1)
                
                left_dots = left_margin_dots + i * (label_size_dots + gap_dots)
                right_dots = left_dots + label_size_dots
                top_dots = top_margin_dots
                bottom_dots = top_dots + label_size_dots
                
                # Draw QR using ImageWin
                dib = ImageWin.Dib(qr_img)
                dib.draw(hdc.GetSafeHdc(), (left_dots, top_dots, right_dots, bottom_dots))
                
                # 3. Draw text label centered below QR code
                text_top = bottom_dots + int(1.0 * mm_to_dot_y) # 1mm below QR code
                
                # Center text under the QR code
                text_width = hdc.GetTextExtent(label_text)[0]
                text_left = left_dots + (label_size_dots - text_width) // 2
                
                hdc.TextOut(text_left, text_top, label_text)
                
            hdc.EndPage()
            
        hdc.EndDoc()
        hdc.DeleteDC()
        return True

def run(server_class=HTTPServer, handler_class=CORSHTTPRequestHandler, port=5001):
    server_address = ('', port)
    httpd = server_class(server_address, handler_class)
    print(f"Starting local Hybrid Print Agent on port {port}...")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    print("Stopping Print Agent.")

if __name__ == '__main__':
    run()
