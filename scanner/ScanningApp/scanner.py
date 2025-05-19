import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client  # <-- Use WIA via pywin32
from PIL import Image
from datetime import datetime
import os   


def scan_documents(paper_size="A4", use_adf=False):
    try:
        wia = win32com.client.Dispatch("WIA.CommonDialog")
        device_manager = win32com.client.Dispatch("WIA.DeviceManager")
        devices = device_manager.DeviceInfos
        if devices.Count == 0:
            messagebox.showerror("Scan Error", "No scanner found.")
            return []
        device = devices.Item(1).Connect()
        items = device.Items
        scanned_files = []
        # Scan one page at a time (WIA doesn't support multi-page by default)
        for i in range(1):
            img = wia.ShowAcquireImage(
                DeviceType=1,  # Scanner
                Intent=0,      # Color
                Bias=0,        # None
                FormatID="{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}",  # BMP
                AlwaysSelectDevice=False,
                UseCommonUI=True if i == 0 else False,
                CancelError=True
            )
            if img is None:
                break
            filename = f"scan_page_{i+1}.bmp"
            img.SaveFile(filename)
            scanned_files.append(filename)
        return scanned_files
    except Exception as e:
        messagebox.showerror("Scan Error", str(e))
        return []

def convert_images_to_pdf(image_paths):
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        images = [Image.open(p).convert("RGB") for p in image_paths]
        output_pdf = f"scanned_{timestamp}.pdf"
        images[0].save(output_pdf, save_all=True, append_images=images[1:])
        return output_pdf
    except Exception as e:
        messagebox.showerror("PDF Error", str(e))
        return None

def upload_files():
    file_paths = filedialog.askopenfilenames(
        title="Select Images",
        filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp")]
    )
    return list(file_paths)

def on_scan():
    paper_size = paper_size_var.get()
    use_adf = adf_var.get()
    status_label.config(text=f"Scanning ({paper_size}, ADF: {use_adf})...")
    window.update_idletasks()

    scanned_images = scan_documents(paper_size, use_adf)
    if scanned_images:
        pdf_path = convert_images_to_pdf(scanned_images)
        if pdf_path:
            status_label.config(text=f"PDF saved: {pdf_path}")
        else:
            status_label.config(text="PDF creation failed.")
    else:
        status_label.config(text="No pages scanned.")

def on_upload():
    files = upload_files()
    if files:
        pdf_path = convert_images_to_pdf(files)
        if pdf_path:
            status_label.config(text=f"Uploaded images saved to PDF: {pdf_path}")
        else:
            status_label.config(text="PDF creation failed.")
    else:
        status_label.config(text="No files selected.")

# === GUI SETUP ===
window = tk.Tk()
window.title("Scanner PDF App")
window.geometry("400x250")

# Paper size dropdown
tk.Label(window, text="Select Paper Size:").pack(pady=(10, 0))
paper_size_var = tk.StringVar(value="A4")
paper_dropdown = ttk.Combobox(window, textvariable=paper_size_var, state="readonly")
paper_dropdown["values"] = ("A4", "A3")
paper_dropdown.pack(pady=5)

# ADF Checkbox
adf_var = tk.BooleanVar()
adf_checkbox = tk.Checkbutton(window, text="Use Document Feeder (ADF)", variable=adf_var)
adf_checkbox.pack(pady=5)

# Buttons
tk.Button(window, text="Scan Document", command=on_scan).pack(pady=10)
tk.Button(window, text="Upload Images", command=on_upload).pack()

# Status
status_label = tk.Label(window, text="", fg="blue", wraplength=350)
status_label.pack(pady=15)

window.mainloop()
