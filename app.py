from flask import Flask, render_template, request, send_file, jsonify
import os
from werkzeug.utils import secure_filename
from converter import DouzoneConverter
import uuid

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 회사명-영문명 매핑
COMPANY_MAP = {
    '베이징그레이스레이저기술유한회사(영업소)': ['Grace Laser Korea Branch'],
    '필립스카본블랙코리아 대표사무소': [
        'PCB KR RO',
        'PHILLIPS CARBON BLACK KOREA REPRESENTATIVE OFFICE'
    ]
}

# 부분일치 검색
def search_company(partial_name):
    results = []
    for kr_name, en_names in COMPANY_MAP.items():
        if partial_name in kr_name:
            results.append({'kr': kr_name, 'en': en_names})
    return results

@app.route('/search_company')
def search():
    query = request.args.get('q', '')
    if len(query) >= 3:
        return jsonify(search_company(query))
    return jsonify([])

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('excel_file')
        english_name = request.form.get('selected_english_name', '').strip()

        if not english_name:
            return "\u274c \uc601\ubb38 \ud68c\uc0ac\uba85\uc744 \uba3c\uc800 \uc120\ud0dd\ud574\uc8fc\uc138\uc694.", 400

        if file and file.filename.endswith(('.xlsx', '.xls')):
            original_filename = secure_filename(file.filename)
            file_root, file_ext = os.path.splitext(original_filename)
            input_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4().hex}_{original_filename}")
            output_filename = f"{file_root}_converted.xlsx"
            output_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4().hex}_{output_filename}")
            file.save(input_path)

            converter = DouzoneConverter()
            success = converter.convert(input_path, output_path, english_name)

            if success:
                return send_file(output_path, as_attachment=True, download_name=output_filename)
            else:
                return "\u274c \ubcc0\ud658 \uc2e4\ud328. \ud30c\uc77c\uc744 \ud655\uc778\ud574\uc8fc\uc138\uc694.", 400

    return render_template('index.html')

if __name__ == '__main__':
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)