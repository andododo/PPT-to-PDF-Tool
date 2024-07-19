from flask import Flask, request, send_file, send_from_directory
import os
import tempfile
from ppt_to_pdf import convert_ppt_to_pdf

app = Flask(__name__, static_folder=".")

@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")

@app.route("/convert", methods=["POST"])
def convert():
    ppt_file = request.files["pptFile"]
    with tempfile.NamedTemporaryFile(delete=False) as tmp_ppt:
        ppt_file.save(tmp_ppt.name)
        ppt_path = tmp_ppt.name

    # get the downloads folder path
    downloads_folder = os.path.expanduser("~/Downloads")
    
    # generate the output PDF file path in the downloads folder
    pdf_file_name = os.path.splitext(ppt_file.filename)[0] + ".pdf"
    pdf_path = os.path.join(downloads_folder, pdf_file_name)
    
    convert_ppt_to_pdf(ppt_path, pdf_path)

    return send_file(pdf_path, as_attachment=True)

if __name__ == "__main__":
    app.run()