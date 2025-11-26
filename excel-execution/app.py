"""
Ứng dụng Flask xử lý import/chỉnh sửa/xuất Excel theo 3 sheet chính.

Cấu trúc tổng quát:
- Cấu hình ứng dụng & thư mục tải lên/cache
- Định nghĩa cột chuẩn cho từng sheet
- Bộ hàm tiện ích xử lý hàng/cột, ngày tháng, tiến độ
- Cơ chế cache JSON: load/save trạng thái `data_store`
- Endpoint giao diện chính và các hành động (import, add/delete row, update cell,
  đổi tên cột, chèn/xóa cột, export)
- Bộ filter Jinja hỗ trợ hiển thị
"""

from flask import Flask, render_template, request, jsonify, abort, send_file
import pandas as pd
from werkzeug.utils import secure_filename
import os
import json
from datetime import datetime, timedelta
import uuid
import requests  

# --- Cảnh báo WhatsApp ---
TWILIO_ACCOUNT_SID = os.environ.get('TWILIO_ACCOUNT_SID', '')
TWILIO_AUTH_TOKEN = os.environ.get('TWILIO_AUTH_TOKEN', '')
TWILIO_WHATSAPP_FROM = os.environ.get('TWILIO_WHATSAPP_FROM', '')  

def _load_phone_recipients():
    if os.path.exists(PHONE_RECIPIENTS_FILE):
        try:
            with open(PHONE_RECIPIENTS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def _send_whatsapp(to_number: str, body: str) -> bool:
    """Gửi tin nhắn WhatsApp qua Twilio. Trả về True nếu thành công."""
    if not (TWILIO_ACCOUNT_SID and TWILIO_AUTH_TOKEN and TWILIO_WHATSAPP_FROM and to_number):
        return False
    url = f"https://api.twilio.com/2010-04-01/Accounts/{TWILIO_ACCOUNT_SID}/Messages.json"
    data = {
        'From': TWILIO_WHATSAPP_FROM,
        'To': to_number,
        'Body': body
    }
    try:
        resp = requests.post(url, data=data, auth=(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN), timeout=10)
        return resp.status_code == 201
    except Exception:
        return False

def _build_whatsapp_body(row: dict) -> str:
    proj = str(row.get('Tên dự án - Mục đích sizing', '')).strip()
    kpi = str(row.get('Thời gian hoàn thành theo KPI', '')).strip()
    status = str(row.get('Tiến độ', '')).strip()
    return (
        f"CẢNH BÁO TIẾN ĐỘ: {status}\n"
        f"Dự án: {proj}\n"
        f"KPI: {kpi}\n"
        f"Vui lòng kiểm tra và xử lý."
    )

def check_and_send_whatsapp_alerts():
    """Quét sheet Sizing, gửi WhatsApp cho các dòng 'Đến hạn' hoặc 'Còn 1 ngày'."""
    _refresh_sizing_progress()
    recipients = _load_phone_recipients()  
    sent = 0
    for row in data_store.get('Sizing', []):
        status = row.get('Tiến độ', '')
        if status in ['Đến hạn', 'Còn 1 ngày']:
            rid = row.get('row_id')
            to_number = recipients.get(rid) or os.environ.get('WHATSAPP_DEFAULT_TO', 'whatsapp:+84847764566')
            if to_number and _send_whatsapp(to_number, _build_whatsapp_body(row)):
                sent += 1
    return sent

"""Khởi tạo ứng dụng và cấu hình chung."""
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024 
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

"""Định nghĩa cột chuẩn cho từng sheet."""
SIZING_COLUMNS = [
    "STT","Mã PYC","Đơn vị","Đầu mối tạo PYC","Đầu mối xử lý","Trạng thái","Thời điểm đẩy yêu cầu","Thời gian hoàn thành theo KPI","Tiến độ","Thời gian hoàn thành ký PNX và đóng y/c","Thời gian ký bản chốt sizing","Tên dự án - Mục đích sizing","Ghi chú"
]
CAP_PHAT_COLUMNS = [
    "STT","Dự án","Đơn vị","Đầu mối y/c","Đầu mối P.HT","Mã SR","Tiến độ, vướng mắc, đề xuất","Thời gian tiếp nhận y/c","Timeline thực hiện theo GNOC","Thời gian hoàn thành","Hoàn thành"
]

CHI_TIET_COLUMNS = [
    "STT","Dự án","Đơn vị","Đầu mối y/c","Đầu mối P.HT","Mã SR","Qúy cấp phát","Số lượng máy chủ","vCPU","Cint","RAM(GB)","SAN(GB)","NAS(GB)","Ceph(GB)","Bigdata(GB)","Archiving(GB)","S3 Object(GB)","Pool/Nguồn tài nguyên","Nhóm tài nguyên","Ghi chú"
]

# Các cột của sheet Chi tiết phải là số, tuyệt đối không parse ngày
CHI_TIET_NUMERIC_COLS = {
    "Số lượng máy chủ","vCPU","Cint","RAM(GB)","SAN(GB)","NAS(GB)",
    "Ceph(GB)","Bigdata(GB)","Archiving(GB)","S3 Object(GB)"
}

"""Tiện ích tạo một hàng trống với `row_id` duy nhất."""
def blank_row(columns):
    row = {col: '' for col in columns}
    row['row_id'] = str(uuid.uuid4())
    return row

def initial_rows(columns, count=5):
    return [blank_row(columns) for _ in range(count)]

"""Bộ nhớ dữ liệu chính trong runtime (3 sheet)."""
data_store = {
    'Sizing': initial_rows(SIZING_COLUMNS),
    'CapPhat': initial_rows(CAP_PHAT_COLUMNS),
    'ChiTiet': initial_rows(CHI_TIET_COLUMNS)
}

"""Đánh lại số thứ tự STT theo vị trí hiện tại."""
def ensure_stt(rows):
    for idx, row in enumerate(rows, start=1):
        row['STT'] = idx

"""Chuẩn hoá dữ liệu hàng theo danh sách cột (loại bỏ NaN/NaT/None)."""
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
        rid = r.get('row_id') or str(uuid.uuid4())
        new_row['row_id'] = rid
        sanitized.append(new_row)
    return sanitized

def _clean_numeric_string(v):
    """Đưa giá trị về chuỗi số đẹp: bỏ .0 nếu là số nguyên; giữ rỗng nếu không phải số."""
    try:
        num = pd.to_numeric(v, errors='coerce')
        if pd.isna(num):
            return ''
        if float(num).is_integer():
            return str(int(num))
        return str(num)
    except Exception:
        return ''

def _fix_chitiet_numeric_rows(rows):
    """Sửa các giá trị bị hiển thị kiểu ngày ở các cột số của sheet Chi tiết.
    Nếu bắt gặp chuỗi dạng dd/mm/yyyy hoặc Timestamp -> chuyển thành rỗng.
    Nếu là số -> chuẩn hoá về chuỗi số.
    """
    import re
    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    for r in rows:
        for c in CHI_TIET_NUMERIC_COLS:
            val = r.get(c, '')
            if val is None:
                r[c] = ''
                continue
            if isinstance(val, (datetime, pd.Timestamp)):
                r[c] = ''
                continue
            s = str(val).strip()
            if not s:
                r[c] = ''
                continue
            if date_re.match(s):
                r[c] = ''
                continue
            # chuẩn hoá số
            r[c] = _clean_numeric_string(s)

ensure_stt(data_store['Sizing'])
ensure_stt(data_store['CapPhat'])
ensure_stt(data_store['ChiTiet'])

"""Thiết lập đường dẫn cache để lưu/khôi phục `data_store`."""
CACHE_DIR = os.path.join(os.path.dirname(__file__), 'cache')
os.makedirs(CACHE_DIR, exist_ok=True)
CACHE_FILE = os.path.join(CACHE_DIR, 'data_store.json')
PHONE_RECIPIENTS_FILE = os.path.join(CACHE_DIR, 'phone_recipients.json')  # mapping row_id -> whatsapp phone

"""Cố gắng parse chuỗi ngày thành Timestamp; lỗi trả về None."""
def _parse_date(val):
    if not val:
        return None
    try:
        dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
        if pd.isna(dt):
            return None
        return dt
    except Exception:
        return None

"""Tính trạng thái tiến độ dựa trên ngày KPI so với hôm nay."""
def _calc_progress_status(kpi_str):
    kpi = _parse_date(kpi_str)
    if not kpi:
        return ""
    today = pd.Timestamp('today').normalize()
    kpi = pd.Timestamp(kpi).normalize()

    if kpi < today:
        return "Quá hạn"

    try:
        wdays = len(pd.bdate_range(start=today, end=kpi, closed='right'))
    except Exception:
        d = today
        wdays = 0
        while d < kpi:
            if d.weekday() < 5:
                wdays += 1
            d += pd.Timedelta(days=1)

    if wdays == 0:
        return "Đến hạn"
    if wdays == 1:
        return "Còn 1 ngày"
    if wdays == 2:
        return "Còn 2 ngày"
    if wdays == 3:
        return "Còn 3 ngày"
    return ""

"""Cập nhật trường 'Tiến độ' của một hàng Sizing."""
def _update_progress_for_row(row):
    row["Tiến độ"] = _calc_progress_status(
        row.get("Thời gian hoàn thành theo KPI", "")
    )

"""Quét toàn bộ sheet Sizing để cập nhật tiến độ."""
def _refresh_sizing_progress():
    try:
        for row in data_store.get('Sizing', []):
            _update_progress_for_row(row)
    except Exception:
        pass

"""Ghi `data_store` ra file JSON cache (best-effort)."""
def save_cache():
    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(data_store, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

"""Đọc cache JSON nếu có và hợp nhất vào `data_store`."""
def load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
            if 'Sizing' in loaded:
                data_store['Sizing'] = sanitize_rows(loaded['Sizing'], SIZING_COLUMNS)
                ensure_stt(data_store['Sizing'])
                for row in data_store['Sizing']:
                    _update_progress_for_row(row)
            if 'CapPhat' in loaded:
                data_store['CapPhat'] = sanitize_rows(loaded['CapPhat'], CAP_PHAT_COLUMNS)
                ensure_stt(data_store['CapPhat'])
            if 'ChiTiet' in loaded:
                data_store['ChiTiet'] = sanitize_rows(loaded['ChiTiet'], CHI_TIET_COLUMNS)
                # Sửa các ô số nếu từng bị lưu dạng ngày (01/01/1970, ...)
                _fix_chitiet_numeric_rows(data_store['ChiTiet'])
                ensure_stt(data_store['ChiTiet'])
        except Exception:
            pass

load_cache()

"""Trang chính: render giao diện với 3 bảng dữ liệu."""
@app.route('/')
def index():
    _refresh_sizing_progress()
    return render_template(
        'index.html',
        sizing_columns=SIZING_COLUMNS,
        cap_phat_columns=CAP_PHAT_COLUMNS,
        chi_tiet_columns=CHI_TIET_COLUMNS,
        sizing_rows=data_store['Sizing'],
        cap_phat_rows=data_store['CapPhat'],
        chi_tiet_rows=data_store['ChiTiet']
    )

"""Chuẩn hoá hiển thị ngày về định dạng dd/mm/YYYY (hoặc rỗng)."""
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

"""Đảm bảo DataFrame có đủ cột, làm sạch và format ngày theo mẫu.

Chú ý: Chỉ định dạng ngày theo TÊN CỘT (keyword) để tránh việc
các cột số (ví dụ trong sheet 'Chi tiết': 'Số lượng máy chủ', 'S3 Object(GB)', ...)
bị hiểu nhầm thành ngày tháng.
"""
def _read_sheet(df: pd.DataFrame, expected_cols):
    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""

    df = df[expected_cols].fillna("").replace({pd.NaT: ""})

    # Chỉ nhận diện cột ngày theo tên cột để tránh nhầm lẫn
    DATE_KEYWORDS = ["Thời", "Timeline", "Qúy"]
    def is_date_col_by_name(name: str) -> bool:
        n = (name or "").lower()
        return any(kw.lower() in n for kw in DATE_KEYWORDS)

    processed = {}
    for col in df.columns:
        series = df[col]
        if is_date_col_by_name(col):
            def fmt(v):
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    return ""
                try:
                    dt = pd.to_datetime(v, dayfirst=True, errors='coerce')
                    if pd.isna(dt):
                        return str(v).strip()
                    return dt.strftime('%d/%m/%Y')
                except Exception:
                    return str(v).strip()
            processed[col] = series.apply(fmt)
        elif col in CHI_TIET_NUMERIC_COLS:
            # Ép về dạng số và xuất chuỗi số; tránh bị hiểu thành ngày
            processed[col] = series.apply(_clean_numeric_string)
        else:
            # Không cố parse ngày ở các cột còn lại; giữ nguyên như chuỗi sạch.
            processed[col] = series.astype(str).apply(lambda x: '' if x.lower() in ['nan','nat'] else x.strip())

    return pd.DataFrame(processed)[expected_cols]

"""Import file Excel: đọc 3 sheet, chuẩn hoá và cập nhật `data_store`."""
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
        _fix_chitiet_numeric_rows(data_store['ChiTiet'])

        _refresh_sizing_progress()

        ensure_stt(data_store['Sizing'])
        ensure_stt(data_store['CapPhat'])
        ensure_stt(data_store['ChiTiet'])
        save_cache()

        return render_template('tables.html', sizing_columns=SIZING_COLUMNS, cap_phat_columns=CAP_PHAT_COLUMNS,
                               chi_tiet_columns=CHI_TIET_COLUMNS,
                               sizing_rows=data_store['Sizing'], cap_phat_rows=data_store['CapPhat'], chi_tiet_rows=data_store['ChiTiet'])
    except Exception as e:
        return jsonify({'error': str(e)}), 500

"""Render partial cho một sheet bất kỳ (HTMX sử dụng)."""
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
    if name == 'Sizing':
        _refresh_sizing_progress()
    rows = data_store[name]
    return render_template('sheet.html', sheet_name=name, sheet_id=sheet_id, columns=columns, rows=rows)

"""Thêm một hàng mới sau vị trí chỉ định trong sheet."""
@app.route('/add-row/<sheet>/<int:after_index>', methods=['POST'])
def add_row(sheet, after_index):
    if sheet not in ['Sizing', 'CapPhat', 'ChiTiet']:
        abort(404)
    columns = SIZING_COLUMNS if sheet == 'Sizing' else (CAP_PHAT_COLUMNS if sheet == 'CapPhat' else CHI_TIET_COLUMNS)
    new_row = {col: '' for col in columns}
    new_row['row_id'] = str(uuid.uuid4())
    target_list = data_store[sheet]
    if after_index < -1 or after_index >= len(target_list):
        target_list.append(new_row)
    else:
        target_list.insert(after_index + 1, new_row)
    ensure_stt(target_list)
    if sheet == 'Sizing':
        _refresh_sizing_progress()
    save_cache()
    sheet_id = 'sizing-sheet' if sheet == 'Sizing' else ('cap-phat-sheet' if sheet == 'CapPhat' else 'chi-tiet-sheet')
    return render_template('sheet.html', sheet_name=sheet, sheet_id=sheet_id, columns=columns, rows=target_list)

"""Xoá một hàng theo chỉ số trong sheet."""
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
    if sheet == 'Sizing':
        _refresh_sizing_progress()
    save_cache()
    columns = SIZING_COLUMNS if sheet == 'Sizing' else (CAP_PHAT_COLUMNS if sheet == 'CapPhat' else CHI_TIET_COLUMNS)
    sheet_id = 'sizing-sheet' if sheet == 'Sizing' else ('cap-phat-sheet' if sheet == 'CapPhat' else 'chi-tiet-sheet')
    return render_template('sheet.html', sheet_name=sheet, sheet_id=sheet_id, columns=columns, rows=target_list)

"""Cập nhật một ô dữ liệu (JSON) và xử lý phụ thuộc tiến độ/KPI."""
@app.route('/update-cell', methods=['POST'])
def update_cell():
    data = request.get_json() or {}
    sheet = data.get('sheet')
    row_index = data.get('row')
    row_id = data.get('rowId')
    col = data.get('col')
    value = data.get('value', '')

    if sheet not in ['Sizing', 'CapPhat', 'ChiTiet']:
        return jsonify({'error': 'Invalid sheet'}), 400
    target_list = data_store[sheet]
    target_row = None
    if row_id:
        for r in target_list:
            if r.get('row_id') == row_id:
                target_row = r
                break
        if target_row is None:
            return jsonify({'error': 'Row not found'}), 400
    else:
        if not isinstance(row_index, int) or row_index < 0 or row_index >= len(target_list):
            return jsonify({'error': 'Invalid row index'}), 400
        target_row = target_list[row_index]
    columns = SIZING_COLUMNS if sheet == 'Sizing' else (CAP_PHAT_COLUMNS if sheet == 'CapPhat' else CHI_TIET_COLUMNS)
    if col not in columns:
        return jsonify({'error': 'Invalid column'}), 400
    target_row[col] = value
    if sheet == 'Sizing' and col == 'Thời gian hoàn thành theo KPI':
        _update_progress_for_row(target_row)
    save_cache()
    return ('', 204)

"""Các filter Jinja hỗ trợ hiển thị rỗng/ngày/đánh class tiến độ."""
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

@app.template_filter('progress_class')
def progress_class(status):
    mapping = {
        'Quá hạn': 'status-overdue',
        'Đến hạn': 'status-due',
        'Còn 1 ngày': 'status-1day',
        'Còn 2 ngày': 'status-2day',
        'Còn 3 ngày': 'status-3day'
    }
    return mapping.get(status, '')

"""Xuất toàn bộ dữ liệu hiện tại ra file Excel (3 sheet)."""
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

"""Truy xuất cặp (rows, columns, sheet_id) theo tên sheet."""
def get_sheet_info(sheet_name):
    if sheet_name == 'Sizing':
        return data_store['Sizing'], SIZING_COLUMNS, 'sizing-sheet'
    elif sheet_name == 'CapPhat':
        return data_store['CapPhat'], CAP_PHAT_COLUMNS, 'cap-phat-sheet'
    elif sheet_name == 'ChiTiet':
        return data_store['ChiTiet'], CHI_TIET_COLUMNS, 'chi-tiet-sheet'
    abort(404)

"""Cập nhật danh sách cột chuẩn tương ứng theo tên sheet."""
def update_columns_constant(sheet_name, new_columns):
    global SIZING_COLUMNS, CAP_PHAT_COLUMNS, CHI_TIET_COLUMNS

    if sheet_name == 'Sizing':
        SIZING_COLUMNS = new_columns
    elif sheet_name == 'CapPhat':
        CAP_PHAT_COLUMNS = new_columns
    elif sheet_name == 'ChiTiet':
        CHI_TIET_COLUMNS = new_columns

"""Đổi tên cột trong một sheet và đồng bộ dữ liệu hàng tương ứng."""
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

"""Chèn/Xoá cột tại vị trí chỉ định, giữ đồng bộ cấu trúc hàng."""
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
            temp_row['row_id'] = row.get('row_id') or str(uuid.uuid4())
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
            if 'row_id' not in row:
                row['row_id'] = str(uuid.uuid4())

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

"""Endpoint thủ công: kích hoạt gửi cảnh báo WhatsApp ngay lập tức."""
@app.route('/trigger-whatsapp-alerts', methods=['POST'])
def trigger_whatsapp_alerts():
    try:
        count = check_and_send_whatsapp_alerts()
        return jsonify({'sent': count}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def _start_whatsapp_daily_scheduler():
    """Tạo thread nền gửi cảnh báo mỗi ngày lúc 08:00."""
    import threading, time

    def loop():
        # Chạy ngay lần đầu để không phải đợi đến 08:00 nếu cần test.
        try:
            check_and_send_whatsapp_alerts()
        except Exception:
            pass
        while True:
            now = datetime.now()
            next_run = now.replace(hour=8, minute=0, second=0, microsecond=0)
            if next_run <= now:
                next_run += timedelta(days=1)
            sleep_seconds = (next_run - now).total_seconds()
            time.sleep(sleep_seconds)
            try:
                check_and_send_whatsapp_alerts()
            except Exception:
                pass

    t = threading.Thread(target=loop, name='whatsapp-scheduler', daemon=True)
    t.start()

"""Điểm vào ứng dụng (chạy development server)."""
if __name__ == '__main__':
    # Khởi động scheduler tự động gửi WhatsApp nếu đã cấu hình Twilio FROM
    if TWILIO_WHATSAPP_FROM:
        _start_whatsapp_daily_scheduler()
    app.run(debug=True)