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
    "STT","Mã PYC","Đơn vị","Đầu mối tạo PYC","Đầu mối xử lý","Trạng thái","Thời điểm đẩy yêu cầu","Thời gian hoàn thành theo KPI","Thời gian hoàn thành ký PNX và đóng y/c","Thời gian ký bản chốt sizing","Tên dự án - Mục đích sizing","Ghi chú"
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

def _format_date(val) -> str:
    if val is None:
        return ""

    if isinstance(val, float) and pd.isna(val):
        return ""

    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%d/%m/%Y")

    try:
        dt = pd.to_datetime(str(val), dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return ""
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return ""

def _read_sheet(df: pd.DataFrame, expected_cols):
    # Bổ sung cột thiếu với giá trị rỗng
    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""

    # Giữ đúng thứ tự cột
    df = df[expected_cols].fillna("").replace({pd.NaT: ""})

    DATE_KEYWORDS = ["Thời", "Timeline", "Qúy"]

    def is_date_column(col_name: str) -> bool:
        for kw in DATE_KEYWORDS:
            if kw.lower() in col_name.lower():
                return True
        return False

    processed = {}
    for col in df.columns:
        series = df[col]
        if pd.api.types.is_datetime64_any_dtype(series) or is_date_column(col):
            # Chỉ format nếu thực sự parse được
            def fmt(v):
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    return ""
                try:
                    # Giữ nguyên nếu là chuỗi không parse được
                    dt = pd.to_datetime(v, dayfirst=True, errors='coerce')
                    if pd.isna(dt):
                        return str(v).strip()
                    return dt.strftime('%d/%m/%Y')
                except Exception:
                    return str(v).strip()
            processed[col] = series.apply(fmt)
        else:
            # Giữ nguyên nội dung text; chỉ strip khoảng trắng
            processed[col] = series.astype(str).apply(lambda x: '' if x.lower() in ['nan','nat'] else x.strip())

    return pd.DataFrame(processed)[expected_cols]

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
        cap_phat_df = pd.read_excel(xl, sheet_name='Cấp phát TN') if 'Cấp phát TN' in xl.sheet_names else pd.DataFrame(columns=CAP_PHAT_COLUMNS)
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

@app.template_filter("format_date")
def format_date_filter(val):
    return _format_date(val)

@app.route('/export')
def export_excel():
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"export_{timestamp}.xlsx"
    path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    try:
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            def build_df(rows, cols):
                if not rows:
                    return pd.DataFrame(columns=cols)
                df = pd.DataFrame(rows)
                DATE_KEYWORDS = ["Thời", "Timeline", "Qúy"]
                def is_date_col(name):
                    return any(kw.lower() in name.lower() for kw in DATE_KEYWORDS)
                for c in cols:
                    if c in df.columns and is_date_col(c):
                        df[c] = df[c].apply(lambda v: _format_date(v) if v not in [None,''] else '')
                return df[cols]

            sizing_df = build_df(data_store['Sizing'], SIZING_COLUMNS)
            cap_df = build_df(data_store['CapPhat'], CAP_PHAT_COLUMNS)
            chi_tiet_df = build_df(data_store['ChiTiet'], CHI_TIET_COLUMNS)

            sizing_df.to_excel(writer, sheet_name='Sizing', index=False)
            cap_df.to_excel(writer, sheet_name='Cấp phát TN', index=False)
            chi_tiet_df.to_excel(writer, sheet_name='Chi tiết', index=False)
        return send_file(path, as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Column execution
def get_sheet_info(sheet_name):
    if sheet_name == 'Sizing':
        return data_store['Sizing'], SIZING_COLUMNS, 'sizing-sheet'
    elif sheet_name == 'CapPhat':
        return data_store['CapPhat'], CAP_PHAT_COLUMNS, 'cap-phat-sheet'
    elif sheet_name == 'ChiTiet':
        return data_store['ChiTiet'], CHI_TIET_COLUMNS, 'chi-tiet-sheet'
    abort(404)

def update_columns_constant(sheet_name, new_columns):
    global SIZING_COLUMNS, CAP_PHAT_COLUMNS, CHI_TIET_COLUMNS

    if sheet_name == 'Sizing':
        SIZING_COLUMNS = new_columns
    elif sheet_name == 'CapPhat':
        CAP_PHAT_COLUMNS = new_columns
    elif sheet_name == 'ChiTiet':
        CHI_TIET_COLUMNS = new_columns

@app.route('/update-col-name', methods=['POST'])
def update_col_name():
    data = request.get_json() or {}
    sheet = data.get('sheet')
    old_col_name = data.get('oldCol')
    new_col_name = data.get('newCol', '').strip()

    if sheet not in data_store or not new_col_name:
        return jsonify({'error': 'Invalid request'}), 400

    rows, columns, sheet_id = get_sheet_info(sheet)

    try:
        col_index = columns.index(old_col_name)
    except ValueError:
        return jsonify({'error': 'Column not found'}), 400

    new_columns = columns[:]
    new_columns[col_index] = new_col_name
    update_columns_constant(sheet, new_columns)

    for row in rows:
        if old_col_name in row:
            row[new_col_name] = row.pop(old_col_name)

    save_cache()
    return ('', 204)

@app.route('/handle-col/<sheet>/<action>/<int:col_index>', methods=['POST'])
def handle_col(sheet, action, col_index):
    if sheet not in data_store:
        abort(404, description=f"Sheet '{sheet}' không tồn tại")

    rows, columns, sheet_id = get_sheet_info(sheet)

    new_col_name = ''
    if action == 'insert':
        new_col_name = request.values.get('new_col_name', '').strip()

        if not new_col_name and request.is_json:
            payload = request.get_json(silent=True) or {}
            new_col_name = payload.get('new_col_name', '').strip()

    if action == 'insert':
        if not new_col_name:
            return jsonify({'error': 'Tên cột không được để trống'}), 400
        if new_col_name in columns:
            return jsonify({'error': f"Tên cột '{new_col_name}' đã tồn tại"}), 400
        if col_index < 0 or col_index > len(columns):
            return jsonify({'error': 'Vị trí cột không hợp lệ'}), 400

        new_columns = columns[:]
        new_columns.insert(col_index, new_col_name)

        update_columns_constant(sheet, new_columns)

        for row in rows:
            temp_row = {col: row.get(col, '') for col in new_columns}
            temp_row[new_col_name] = ''
            row.clear()
            row.update(temp_row)

        return render_template(
            'tables.html',
            sizing_columns=SIZING_COLUMNS,
            sizing_rows=data_store['Sizing'],
            cap_phat_columns=CAP_PHAT_COLUMNS,
            cap_phat_rows=data_store['CapPhat'],
            chi_tiet_columns=CHI_TIET_COLUMNS,
            chi_tiet_rows=data_store['ChiTiet']
        ), 200

    if action == 'delete':
        if col_index < 0 or col_index >= len(columns):
            return jsonify({'error': 'Vị trí cột không hợp lệ'}), 400
        col_name_to_delete = columns[col_index]

        if col_name_to_delete == 'STT':
            return jsonify({'error': 'Không thể xóa cột STT'}), 400

        new_columns = columns[:]
        del new_columns[col_index]

        for row in rows:
            row.pop(col_name_to_delete, None)

        update_columns_constant(sheet, new_columns)

        return render_template(
            'tables.html',
            sizing_columns=SIZING_COLUMNS,
            sizing_rows=data_store['Sizing'],
            cap_phat_columns=CAP_PHAT_COLUMNS,
            cap_phat_rows=data_store['CapPhat'],
            chi_tiet_columns=CHI_TIET_COLUMNS,
            chi_tiet_rows=data_store['ChiTiet']
        ), 200

    return jsonify({'error': f'Hành động {action} không được hỗ trợ'}), 400

if __name__ == '__main__':
    app.run(debug=True)