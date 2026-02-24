import os
import zipfile
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
from io import BytesIO

app = Flask(__name__, template_folder='templates', static_folder='static')

def replace_placeholders(doc, data_row):
    """Inabadilisha tagi [TAGI] kwenda thamani halisi."""
    data_dict = {str(k).upper(): str(v) for k, v in data_row.items()}
    
    def process_text(text):
        for key, value in data_dict.items():
            tag = f"[{key}]"
            if tag in text:
                text = text.replace(tag, value)
        return text

    for paragraph in doc.paragraphs:
        paragraph.text = process_text(paragraph.text)
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.text = process_text(paragraph.text)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_files():
    try:
        word_file = request.files.get('word_template')
        excel_file = request.files.get('excel_data')
        filename_col = request.form.get('filename_column', '').strip()

        if not word_file or not excel_file or not filename_col:
            return jsonify({"error": "Tafadhali kamilisha fomu na pakia faili zote."}), 400

        try:
            df = pd.read_excel(excel_file)
            df.columns = [str(c).upper().strip() for c in df.columns]
        except Exception:
            return jsonify({"error": "Faili la Excel halisomeki. Tumia .xlsx ya kisasa."}), 400

        col_search = filename_col.upper().strip()
        if col_search not in df.columns:
            return jsonify({"error": f"Safu '{filename_col}' haiko kwenye Excel. Safu zilizopo ni: {', '.join(df.columns)}"}), 400

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED) as zip_file:
            for index, row in df.iterrows():
                word_file.seek(0)
                doc = Document(word_file)
                replace_placeholders(doc, row.to_dict())
                
                doc_io = BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
                file_name_val = str(row.get(col_search, index)).replace("/", "-").replace("\\", "-")
                zip_file.writestr(f"Barua_{file_name_val}.docx", doc_io.getvalue())

        zip_buffer.seek(0)
        return send_file(
            zip_buffer, 
            mimetype='application/zip', 
            as_attachment=True, 
            download_name='Ngosha_Bulk_Letters.zip'
        )

    except Exception as e:
        return jsonify({"error": f"Kosa la Seva: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True)
