from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
from io import BytesIO
from pdf2docx import Converter

app = Flask(__name__)

@app.route('/')
def upload_file():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return "No file uploaded", 400
    
    pdf_file = request.files['file']
    if pdf_file.filename == '':
        return "No selected file", 400
    
    if pdf_file and pdf_file.filename.endswith('.pdf'):
        # Secure the filename and read the file into BytesIO
        filename = secure_filename(pdf_file.filename)
        pdf_bytes = pdf_file.read()

        # Save the PDF to a temporary file because pdf2docx works best with file paths
        with open("temp.pdf", "wb") as temp_pdf:
            temp_pdf.write(pdf_bytes)

        # Convert PDF to Word
        docx_io = BytesIO()
        converter = Converter("temp.pdf")  # Use the file path here
        converter.convert(docx_io)  # Pass the BytesIO for the DOCX output
        converter.close()
        docx_io.seek(0)

        # Clean up the temporary PDF file
        import os
        os.remove("temp.pdf")

        # Send the Word file as a downloadable response
        return send_file(
            docx_io, 
            as_attachment=True, 
            download_name=filename.replace('.pdf', '.docx'),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        return "Invalid file format. Please upload a PDF.", 400

if __name__ == '__main__':
    app.run(debug=True)
