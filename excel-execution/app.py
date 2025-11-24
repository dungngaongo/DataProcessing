from flask import Flask, render_template, request, jsonify, abort, send_file
import pandas as pd
from werkzeug.utils import secure_filename
import os
import json
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB limit
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

SIZING_COLUMNS = [
    "STT","Mã PYC","Đơn vị","Đầu mối tạo PYC","Đầu mối xử lý","Status","Thời điểm đẩy y/c","Thời gian hoàn thành theo KPI","Thời gian hoàn thành ký PNX và đóng y/c","Thời gian ký bản chốt sizing","Tên dự án - Mục đích sizing","Ghi chú"
]
CAP_PHAT_COLUMNS = [
    "STT","Dự án","Đơn vị","Đầu mối y/c","Đầu mối P.HT","Mã SR","Tiến độ, vướng mắc, đề xuất","Thời gian tiếp nhận y/c","Timeline thực hiện theo GNOC","Thời gian hoàn thành","Hoàn thành"
]

CHI_TIET_COLUMNS = [
    "STT","Dự án","Đơn vị","Đầu mối y/c","Đầu mối P.HT","Mã SR","Qúy cấp phát","Số lượng máy chủ","vCPU","Cint","RAM(GB)","SAN(GB)","NAS(GB)","Ceph(GB)","Bigdata(GB)","Archiving(GB)","S3 Object(GB)","Pool/Nguồn tài nguyên","Nhóm tài nguyên","Ghi chú"
]

def blank_row(columns):
    return {col: '' for col in columns}

def initial_rows(columns, count=5):
    return [blank_row(columns) for _ in range(count)]

data_store = {
    'Sizing': initial_rows(SIZING_COLUMNS),
    'CapPhat': initial_rows(CAP_PHAT_COLUMNS),
    'ChiTiet': initial_rows(CHI_TIET_COLUMNS)
}

def ensure_stt(rows):
    for idx, row in enumerate(rows, start=1):
        row['STT'] = idx

def sanitize_rows(rows, columns):
    sanitized = []
    for r in rows:
        new_row = {}
        for c in columns:
            val = r.get(c, '')
            if val is None or (isinstance(val, float) and pd.isna(val)):
                val = ''
            else:
                sval = str(val).strip()
                if sval.lower() in ['nan', 'nat']:
                    val = ''
            new_row[c] = val
        sanitized.append(new_row)
    return sanitized

ensure_stt(data_store['Sizing'])
ensure_stt(data_store['CapPhat'])
ensure_stt(data_store['ChiTiet'])

CACHE_DIR = os.path.join(os.path.dirname(__file__), 'cache')
os.makedirs(CACHE_DIR, exist_ok=True)
CACHE_FILE = os.path.join(CACHE_DIR, 'data_store.json')

def save_cache():
    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(data_store, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
            if 'Sizing' in loaded:
                data_store['Sizing'] = sanitize_rows(loaded['Sizing'], SIZING_COLUMNS)
                ensure_stt(data_store['Sizing'])
            if 'CapPhat' in loaded:
                data_store['CapPhat'] = sanitize_rows(loaded['CapPhat'], CAP_PHAT_COLUMNS)
                ensure_stt(data_store['CapPhat'])
            if 'ChiTiet' in loaded:
                data_store['ChiTiet'] = sanitize_rows(loaded['ChiTiet'], CHI_TIET_COLUMNS)
                ensure_stt(data_store['ChiTiet'])
        except Exception:
            pass

load_cache()

@app.route('/')
def index():
    return render_template(
        'index.html',
        sizing_columns=SIZING_COLUMNS,
        cap_phat_columns=CAP_PHAT_COLUMNS,
        chi_tiet_columns=CHI_TIET_COLUMNS,
        sizing_rows=data_store['Sizing'],
        cap_phat_rows=data_store['CapPhat'],
        chi_tiet_rows=data_store['ChiTiet']
    )


def _read_sheet(df: pd.DataFrame, expected_cols):
    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""
    df = df[expected_cols]
    df = df.fillna("").replace({pd.NaT: ""})
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime('%Y-%m-%d').replace('NaT', '')
    return df

@app.route('/import', methods=['POST'])
def import_excel():
    file = request.files.get('excel_file')
    if not file:
        return jsonify({'error': 'No file provided'}), 400

    filename = secure_filename(file.filename)
    path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(path)

    try:
        xl = pd.ExcelFile(path)
        sizing_df = pd.read_excel(xl, sheet_name='Sizing') if 'Sizing' in xl.sheet_names else pd.DataFrame(columns=SIZING_COLUMNS)
        cap_phat_df = pd.read_excel(xl, sheet_name='Cấp phát tài nguyên') if 'Cấp phát tài nguyên' in xl.sheet_names else pd.DataFrame(columns=CAP_PHAT_COLUMNS)
        chi_tiet_df = pd.read_excel(xl, sheet_name='Chi tiết') if 'Chi tiết' in xl.sheet_names else pd.DataFrame(columns=CHI_TIET_COLUMNS)

        sizing_df = _read_sheet(sizing_df, SIZING_COLUMNS)
        cap_phat_df = _read_sheet(cap_phat_df, CAP_PHAT_COLUMNS)
        chi_tiet_df = _read_sheet(chi_tiet_df, CHI_TIET_COLUMNS)

        data_store['Sizing'] = sanitize_rows(sizing_df.to_dict(orient='records'), SIZING_COLUMNS)
        data_store['CapPhat'] = sanitize_rows(cap_phat_df.to_dict(orient='records'), CAP_PHAT_COLUMNS)
        data_store['ChiTiet'] = sanitize_rows(chi_tiet_df.to_dict(orient='records'), CHI_TIET_COLUMNS)
        ensure_stt(data_store['Sizing'])
        ensure_stt(data_store['CapPhat'])
        ensure_stt(data_store['ChiTiet'])
        save_cache()

        return render_template('tables.html', sizing_columns=SIZING_COLUMNS, cap_phat_columns=CAP_PHAT_COLUMNS,
                               chi_tiet_columns=CHI_TIET_COLUMNS,
                               sizing_rows=data_store['Sizing'], cap_phat_rows=data_store['CapPhat'], chi_tiet_rows=data_store['ChiTiet'])
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/sheet/<name>')
def sheet(name):
    if name not in ['Sizing', 'CapPhat', 'ChiTiet']:
        abort(404)
    if name == 'Sizing':
        columns = SIZING_COLUMNS
        sheet_id = 'sizing-sheet'
    elif name == 'CapPhat':
        columns = CAP_PHAT_COLUMNS
        sheet_id = 'cap-phat-sheet'
    else:
        columns = CHI_TIET_COLUMNS
        sheet_id = 'chi-tiet-sheet'
    rows = data_store[name]
    return render_template('sheet.html', sheet_name=name, sheet_id=sheet_id, columns=columns, rows=rows)

@app.route('/add-row/<sheet>/<int:after_index>', methods=['POST'])
def add_row(sheet, after_index):
    if sheet not in ['Sizing', 'CapPhat', 'ChiTiet']:
        abort(404)
    columns = SIZING_COLUMNS if sheet == 'Sizing' else (CAP_PHAT_COLUMNS if sheet == 'CapPhat' else CHI_TIET_COLUMNS)
    new_row = {col: '' for col in columns}
    target_list = data_store[sheet]
    if after_index < -1 or after_index >= len(target_list):
        target_list.append(new_row)
    else:
        target_list.insert(after_index + 1, new_row)
    ensure_stt(target_list)
    save_cache()
    sheet_id = 'sizing-sheet' if sheet == 'Sizing' else ('cap-phat-sheet' if sheet == 'CapPhat' else 'chi-tiet-sheet')
    return render_template('sheet.html', sheet_name=sheet, sheet_id=sheet_id, columns=columns, rows=target_list)

@app.route('/delete-row/<sheet>/<int:row_index>', methods=['POST'])
def delete_row(sheet, row_index):
    if sheet not in ['Sizing', 'CapPhat', 'ChiTiet']:
        abort(404)
    target_list = data_store[sheet]
    if row_index < 0 or row_index >= len(target_list):
        return jsonify({'error': 'Invalid row index'}), 400
    target_list.pop(row_index)
    if target_list:
        ensure_stt(target_list)
    save_cache()
    columns = SIZING_COLUMNS if sheet == 'Sizing' else (CAP_PHAT_COLUMNS if sheet == 'CapPhat' else CHI_TIET_COLUMNS)
    sheet_id = 'sizing-sheet' if sheet == 'Sizing' else ('cap-phat-sheet' if sheet == 'CapPhat' else 'chi-tiet-sheet')
    return render_template('sheet.html', sheet_name=sheet, sheet_id=sheet_id, columns=columns, rows=target_list)

@app.route('/update-cell', methods=['POST'])
def update_cell():
    data = request.get_json() or {}
    sheet = data.get('sheet')
    row_index = data.get('row')
    col = data.get('col')
    value = data.get('value', '')

    if sheet not in ['Sizing', 'CapPhat', 'ChiTiet']:
        return jsonify({'error': 'Invalid sheet'}), 400
    target_list = data_store[sheet]
    if not isinstance(row_index, int) or row_index < 0 or row_index >= len(target_list):
        return jsonify({'error': 'Invalid row index'}), 400
    columns = SIZING_COLUMNS if sheet == 'Sizing' else (CAP_PHAT_COLUMNS if sheet == 'CapPhat' else CHI_TIET_COLUMNS)
    if col not in columns:
        return jsonify({'error': 'Invalid column'}), 400
    target_list[row_index][col] = value
    if col != 'STT':  
        ensure_stt(target_list)
    save_cache()
    return ('', 204)

@app.template_filter('blanknan')
def blanknan(val):
    if val is None:
        return ''
    if isinstance(val, float) and pd.isna(val):
        return ''
    sval = str(val).strip()
    if sval.lower() in ['nan', 'nat']:
        return ''
    return val

@app.route('/export')
def export_excel():
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"export_{timestamp}.xlsx"
    path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    try:
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            sizing_df = pd.DataFrame(data_store['Sizing'])[SIZING_COLUMNS] if data_store['Sizing'] else pd.DataFrame(columns=SIZING_COLUMNS)
            cap_df = pd.DataFrame(data_store['CapPhat'])[CAP_PHAT_COLUMNS] if data_store['CapPhat'] else pd.DataFrame(columns=CAP_PHAT_COLUMNS)
            chi_tiet_df = pd.DataFrame(data_store['ChiTiet'])[CHI_TIET_COLUMNS] if data_store['ChiTiet'] else pd.DataFrame(columns=CHI_TIET_COLUMNS)
            sizing_df.to_excel(writer, sheet_name='Sizing', index=False)
            cap_df.to_excel(writer, sheet_name='Cấp phát tài nguyên', index=False)
            chi_tiet_df.to_excel(writer, sheet_name='Chi tiết', index=False)
        return send_file(path, as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True)
