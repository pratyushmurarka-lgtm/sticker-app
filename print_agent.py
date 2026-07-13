import os
import sys
import json
import base64
import tempfile
import io
import ctypes
from ctypes import wintypes
from http.server import HTTPServer, BaseHTTPRequestHandler
import win32print
import win32ui
import win32con
import win32api
from PIL import Image, ImageWin
import segno

# ---------------------------------------------------------
# Windows Common Control Print Dialog Structures & APIs
# ---------------------------------------------------------
comdlg32 = ctypes.windll.comdlg32
kernel32 = ctypes.windll.kernel32

PD_ALLPAGES = 0x00000000
PD_SELECTION = 0x00000001
PD_PAGENUMS = 0x00000002
PD_NOSELECTION = 0x00000004
PD_NOPAGENUMS = 0x00000008
PD_RETURNDC = 0x00000100
PD_USEDEVMODECOPIESANDCOLLATE = 0x00040000

class PRINTDLGW(ctypes.Structure):
    _fields_ = [
        ("lStructSize", wintypes.DWORD),
        ("hwndOwner", wintypes.HWND),
        ("hDevMode", wintypes.HANDLE),
        ("hDevNames", wintypes.HANDLE),
        ("hDC", wintypes.HDC),
        ("Flags", wintypes.DWORD),
        ("nFromPage", wintypes.WORD),
        ("nToPage", wintypes.WORD),
        ("nMinPage", wintypes.WORD),
        ("nMaxPage", wintypes.WORD),
        ("nCopies", wintypes.WORD),
        ("hInstance", wintypes.HINSTANCE),
        ("lCustData", wintypes.LPARAM),
        ("lpfnPrintHook", ctypes.c_void_p),
        ("lpfnSetupHook", ctypes.c_void_p),
        ("lpPrintTemplateName", wintypes.LPCWSTR),
        ("lpSetupTemplateName", wintypes.LPCWSTR),
        ("hPrintTemplate", wintypes.HANDLE),
        ("hSetupTemplate", wintypes.HANDLE),
    ]

comdlg32.PrintDlgW.argtypes = [ctypes.POINTER(PRINTDLGW)]
comdlg32.PrintDlgW.restype = wintypes.BOOL

GlobalLock = kernel32.GlobalLock
GlobalLock.argtypes = [wintypes.HANDLE]
GlobalLock.restype = ctypes.c_void_p

GlobalUnlock = kernel32.GlobalUnlock
GlobalUnlock.argtypes = [wintypes.HANDLE]
GlobalUnlock.restype = wintypes.BOOL

GlobalFree = kernel32.GlobalFree
GlobalFree.argtypes = [wintypes.HANDLE]
GlobalFree.restype = wintypes.HANDLE

def show_print_dialog(max_pages=9999):
    """
    Displays the standard Windows Common Print Dialog.
    Allows the user to select the printer, copies, and page range.
    """
    pd = PRINTDLGW()
    pd.lStructSize = ctypes.sizeof(PRINTDLGW)
    pd.hwndOwner = 0
    pd.Flags = PD_RETURNDC | PD_USEDEVMODECOPIESANDCOLLATE
    pd.nMinPage = 1
    pd.nMaxPage = max_pages
    pd.nFromPage = 1
    pd.nToPage = max_pages
    
    if comdlg32.PrintDlgW(ctypes.byref(pd)):
        # Extract selected printer name from hDevNames structure
        printer_name = ""
        if pd.hDevNames:
            ptr = GlobalLock(pd.hDevNames)
            if ptr:
                try:
                    # wDeviceOffset is located at byte offset 2 inside the DEVNAMES structure
                    wDeviceOffset = ctypes.cast(ptr + 2, ctypes.POINTER(ctypes.c_ushort)).contents.value
                    printer_name = ctypes.wstring_at(ptr + wDeviceOffset * 2)
                finally:
                    GlobalUnlock(pd.hDevNames)
                    
        from_page = pd.nFromPage
        to_page = pd.nToPage
        copies = pd.nCopies
        
        # Check if user selected custom page numbers range
        is_range = bool(pd.Flags & PD_PAGENUMS)
        
        # Clean up handles allocated by PrintDlg to prevent memory leaks
        if pd.hDevMode:
            GlobalFree(pd.hDevMode)
        if pd.hDevNames:
            GlobalFree(pd.hDevNames)
        if pd.hDC:
            ctypes.windll.gdi32.DeleteDC(pd.hDC)
            
        return {
            "success": True,
            "printer_name": printer_name,
            "from_page": from_page,
            "to_page": to_page,
            "is_range": is_range,
            "copies": copies
        }
    else:
        return {
            "success": False
        }


# ---------------------------------------------------------
# Print Agent Server Handler
# ---------------------------------------------------------
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

    def print_pdf(self, pdf_base64, default_printer):
        pdf_bytes = base64.b64decode(pdf_base64)
        fd, temp_pdf_path = tempfile.mkstemp(suffix=".pdf")
        try:
            with os.fdopen(fd, "wb") as f:
                f.write(pdf_bytes)
                
            # Get actual page count using pypdf
            from pypdf import PdfReader, PdfWriter
            try:
                reader = PdfReader(temp_pdf_path)
                max_pages = len(reader.pages)
            except Exception:
                max_pages = 9999
                
            # Show Windows Print Dialog
            dlg_res = show_print_dialog(max_pages)
            if not dlg_res["success"]:
                raise RuntimeError("Print job cancelled by operator.")
                
            selected_printer = dlg_res["printer_name"] or default_printer
            from_page = dlg_res["from_page"]
            to_page = dlg_res["to_page"]
            is_range = dlg_res["is_range"]
            copies = dlg_res["copies"]
            
            print_file_path = temp_pdf_path
            
            # If user specified a custom range, split PDF dynamically
            if is_range and max_pages != 9999:
                fd_split, split_pdf_path = tempfile.mkstemp(suffix=".pdf")
                os.close(fd_split)
                try:
                    writer = PdfWriter()
                    start = max(0, from_page - 1)
                    end = min(max_pages, to_page)
                    for idx in range(start, end):
                        writer.add_page(reader.pages[idx])
                    with open(split_pdf_path, "wb") as out_f:
                        writer.write(out_f)
                    print_file_path = split_pdf_path
                except Exception as split_err:
                    print(f"Failed to split range: {split_err}")
            
            # Set default printer dynamically
            original_printer = win32print.GetDefaultPrinter()
            if original_printer != selected_printer:
                win32print.SetDefaultPrinter(selected_printer)
                
            # Spool the print job (loop copies times)
            for _ in range(max(1, copies)):
                win32api.ShellExecute(0, "print", print_file_path, None, ".", 0)
                win32api.Sleep(500)
                
            # Revert default printer if changed
            if original_printer != selected_printer:
                win32api.Sleep(2000)
                win32print.SetDefaultPrinter(original_printer)
                
            # Schedule split PDF file cleanup if created
            if print_file_path != temp_pdf_path:
                def cleanup_split():
                    try:
                        win32api.Sleep(5000)
                        os.remove(split_pdf_path)
                    except:
                        pass
                import threading
                threading.Thread(target=cleanup_split).start()
                
            return True
        finally:
            def cleanup():
                try:
                    win32api.Sleep(5000)
                    os.remove(temp_pdf_path)
                except:
                    pass
            import threading
            threading.Thread(target=cleanup).start()

    def print_sequential(self, qrs, default_printer, label_size_mm, gap_mm, top_margin_mm):
        total_qrs = len(qrs)
        if total_qrs == 0:
            return True
            
        labels_per_page = 4
        max_pages = (total_qrs + labels_per_page - 1) // labels_per_page
        
        # Show Windows Print Dialog
        dlg_res = show_print_dialog(max_pages)
        if not dlg_res["success"]:
            raise RuntimeError("Print job cancelled by operator.")
            
        selected_printer = dlg_res["printer_name"] or default_printer
        from_page = dlg_res["from_page"]
        to_page = dlg_res["to_page"]
        is_range = dlg_res["is_range"]
        copies = dlg_res["copies"]
        
        # If user specified a page range, slice QRs list
        if is_range:
            start_idx = max(0, (from_page - 1) * labels_per_page)
            end_idx = min(total_qrs, to_page * labels_per_page)
            qrs = qrs[start_idx:end_idx]
            total_qrs = len(qrs)
            if total_qrs == 0:
                return True
            max_pages = (total_qrs + labels_per_page - 1) // labels_per_page
            
        # Loop copies times to print the batch
        for copy_idx in range(max(1, copies)):
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(selected_printer)
            
            dpi_x = hdc.GetDeviceCaps(win32con.LOGPIXELSX)
            dpi_y = hdc.GetDeviceCaps(win32con.LOGPIXELSY)
            
            mm_to_dot_x = dpi_x / 25.4
            mm_to_dot_y = dpi_y / 25.4
            
            label_size_dots = int(label_size_mm * mm_to_dot_x)
            gap_dots = int(gap_mm * mm_to_dot_x)
            top_margin_dots = int(top_margin_mm * mm_to_dot_y)
            left_margin_dots = int(3.0 * mm_to_dot_x)
            
            hdc.StartDoc("Sequential QR Labels")
            
            for p in range(max_pages):
                hdc.StartPage()
                
                font_height = int(3.0 * mm_to_dot_y)
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
                    
                    parts = qr_value.split("|")
                    if len(parts) >= 7:
                        label_text = f"{parts[2].strip()}-{parts[6].strip()}"
                    else:
                        label_text = qr_value
                        
                    qr = segno.make(qr_value, error='M')
                    buf = io.BytesIO()
                    qr.save(buf, kind='png', scale=10, border=1)
                    buf.seek(0)
                    qr_img = Image.open(buf)
                    
                    left_dots = left_margin_dots + i * (label_size_dots + gap_dots)
                    right_dots = left_dots + label_size_dots
                    top_dots = top_margin_dots
                    bottom_dots = top_dots + label_size_dots
                    
                    dib = ImageWin.Dib(qr_img)
                    dib.draw(hdc.GetSafeHdc(), (left_dots, top_dots, right_dots, bottom_dots))
                    
                    text_top = bottom_dots + int(1.0 * mm_to_dot_y)
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
