import os
from tkinter import Tk, filedialog, messagebox
from PIL import Image
from datetime import datetime
import win32com.client

# === SCANNING FUNCTION ===
def scan_and_save(scan_output="scanned.bmp"):
    try:
        wia = win32com.client.Dispatch("WIA.CommonDialog")
        scanner = wia.ShowSelectDevice()
        if not scanner:
            print("No scanner selected.")
            return None

        item = scanner.Items[0]
        image = wia.ShowTransfer(item, "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}")  # BMP format
        image.SaveFile(scan_output)
        return scan_output
    except Exception as e:
        print("Scanning error:", e)
        return None

# === CONVERT IMAGE TO PDF ===
def convert_image_to_pdf(image_paths, output_pdf):
    images = []
    for img_path in image_paths:
        img = Image.open(img_path).convert("RGB")
        images.append(img)
    images[0].save(output_pdf, save_all=True, append_images=images[1:])
    print(f"PDF saved to: {output_pdf}")

# === FILE UPLOAD FUNCTION ===
def upload_files():
    root = Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="Select Image Files", filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp")]
    )
    return list(file_paths)

# === MAIN INTERFACE ===
def main():
    print("Choose an option:")
    print("1. Scan document")
    print("2. Upload images")
    choice = input("Enter 1 or 2: ")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    pdf_output = f"scanned_{timestamp}.pdf"

    if choice == "1":
        scanned_file = scan_and_save()
        if scanned_file:
            convert_image_to_pdf([scanned_file], pdf_output)
    elif choice == "2":
        uploaded_files = upload_files()
        if uploaded_files:
            convert_image_to_pdf(uploaded_files, pdf_output)
        else:
            print("No files selected.")
    else:
        print("Invalid choice.")

if __name__ == "__main__":
    main()
