from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import win32com.client
from PIL import Image
from datetime import datetime
import os

app = Flask(__name__)
app.secret_key = "scanner_secret"

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def scan_documents():
    try:
        wia = win32com.client.Dispatch("WIA.CommonDialog")
        device_manager = win32com.client.Dispatch("WIA.DeviceManager")
        devices = device_manager.DeviceInfos
        if devices.Count == 0:
            return None, "No scanner found."
        device = devices.Item(1).Connect()
        img = wia.ShowAcquireImage(
            DeviceType=1, Intent=0, Bias=0,
            FormatID="{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}",
            AlwaysSelectDevice=False, UseCommonUI=True, CancelError=True
        )
        if img is None:
            return None, "Scan cancelled."
        filename = os.path.join(UPLOAD_FOLDER, "scan_page_1.bmp")
        img.SaveFile(filename)
        return filename, None
    except Exception as e:
        return None, str(e)

def convert_images_to_pdf(image_paths):
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        images = [Image.open(p).convert("RGB") for p in image_paths]
        output_pdf_name = f"scanned_{timestamp}.pdf"
        output_pdf = os.path.join(UPLOAD_FOLDER, output_pdf_name)
        images[0].save(output_pdf, save_all=True, append_images=images[1:])
        return output_pdf
    except Exception as e:
        return None

@app.route("/", methods=["GET", "POST"])
def index():
    pdf_path = None
    if request.method == "POST":
        if "scan" in request.form:
            img_path, error = scan_documents()
            if error:
                flash(error, "danger")
            elif img_path:
                pdf_path = convert_images_to_pdf([img_path])
                if pdf_path:
                    flash("Scan complete. Download your PDF below.", "success")
        elif "upload" in request.form:
            files = request.files.getlist("images")
            paths = []
            for f in files:
                if f.filename:
                    path = os.path.join(UPLOAD_FOLDER, f.filename)
                    f.save(path)
                    paths.append(path)
            if paths:
                pdf_path = convert_images_to_pdf(paths)
                if pdf_path:
                    flash("Images uploaded. Download your PDF below.", "success")
    return render_template("index.html", pdf_path=pdf_path)

@app.route("/download/<filename>")
def download(filename):
    return send_file(os.path.join(UPLOAD_FOLDER, filename), as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)