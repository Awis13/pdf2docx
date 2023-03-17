from flask import Flask, request, send_file
from werkzeug.utils import secure_filename
from main import pdf_to_docx
import os

app = Flask(__name__)

@app.route('/')
def index():
    return '''
    <!doctype html>
    <html>
      <head>
        <title>PDF to DOCX Converter</title>
      </head>
      <body>
        <h1>PDF to DOCX Converter</h1>
        <form method="POST" action="/convert" enctype="multipart/form-data">
          <input type="file" name="pdf_file" accept=".pdf">
          <button type="submit">Convert</button>
        </form>
      </body>
    </html>
    '''

@app.route('/convert', methods=['POST'])
def convert():
    pdf_file = request.files['pdf_file']
    filename = secure_filename(pdf_file.filename)
    input_pdf = os.path.join("uploads", filename)
    output_docx = os.path.splitext(input_pdf)[0] + ".docx"

    pdf_file.save(input_pdf)

    # Replace main(input_pdf, output_docx) with the line below
    pdf_to_docx(input_pdf, output_docx)

    return send_file(output_docx, as_attachment=True)

if __name__ == '__main__':
    if not os.path.exists("uploads"):
        os.makedirs("uploads")
    app.run(debug=True)
