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
from typing import Optional, List
import pandas as pd
from werkzeug.utils import secure_filename
import os
try:
    # Tự động nạp biến môi trường từ file .env nếu có
    from dotenv import load_dotenv
    # Ưu tiên tìm .env trong thư mục dự án hiện tại (excel-execution)
    _ENV_PATH = os.path.join(os.path.dirname(__file__), '.env')
    load_dotenv(_ENV_PATH)
except Exception:
    # Không bắt buộc phải có python-dotenv; nếu thiếu sẽ dùng biến môi trường hệ thống
    pass
import json
from datetime import datetime, timedelta
import uuid
import requests  

# --- Cảnh báo WhatsApp ---
TWILIO_ACCOUNT_SID = os.environ.get('TWILIO_ACCOUNT_SID', '')
TWILIO_AUTH_TOKEN = os.environ.get('TWILIO_AUTH_TOKEN', '')
TWILIO_WHATSAPP_FROM = os.environ.get('TWILIO_WHATSAPP_FROM', '')  
TWILIO_CONTENT_SID = os.environ.get('TWILIO_CONTENT_SID', '')  

# Mapping cố định giữa "đầu mối" và số WhatsApp
OWNER_PHONE_MAP = {
    "thongnv31": "whatsapp:+84333629091",
    "ductn8": "whatsapp:+84335371306",
    "khanhnd23": "whatsapp:+84383522722",
    "vinhtq18": "whatsapp:+84968468868",
    "haipn": "whatsapp:+84962422102",
    "dungnt": "whatsapp:+84847764566",
}

ALWAYS_NOTIFY = ["thongnv31", "haipn", "dungnt"]
ALWAYS_NOTIFY_NUMBERS = [OWNER_PHONE_MAP[k] for k in ALWAYS_NOTIFY if OWNER_PHONE_MAP.get(k)]

def _prepare_message_for_recipient(sheet_name: str, row: dict, to_number: str, base_body: str, variables: Optional[dict]):
    if variables is None:
        variables = {}
    body = base_body
    try:
        if to_number in ALWAYS_NOTIFY_NUMBERS:
            if sheet_name == 'Sizing':
                owner_label = 'Đầu mối xử lý'
                owner_value = str(row.get('Đầu mối xử lý', '')).strip()
            elif sheet_name == 'CapPhat':
                owner_label = 'Đầu mối P.HT'
                owner_value = str(row.get('Đầu mối P.HT', '')).strip()
            else:
                owner_label = 'Đầu mối'
                owner_value = ''
            if owner_value:
                body = f"{base_body}\nĐầu mối phụ trách: {owner_value}"
            # Bổ sung biến cho template (gộp sẵn câu để template dùng trực tiếp nếu cần)
            variables = {
                **variables,
                "owner_label": owner_label,
                "owner": owner_value,
                "owner_supervisor": (f"Đầu mối phụ trách: {owner_value}" if owner_value else "")
            }
    except Exception:
        pass
    return body, variables

def _load_phone_recipients():
    if os.path.exists(PHONE_RECIPIENTS_FILE):
        try:
            with open(PHONE_RECIPIENTS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def _send_whatsapp(to_number: str, body: str, variables: Optional[dict] = None) -> bool:
    """Gửi tin nhắn WhatsApp qua Twilio. Trả về True nếu thành công.

    Nếu có `TWILIO_CONTENT_SID`, sẽ gửi qua Content API (template) để tránh lỗi 63016 (ngoài 24h window).
    `variables` là dict cho ContentVariables (JSON string) nếu dùng template có placeholders.
    """
    if not (TWILIO_ACCOUNT_SID and TWILIO_AUTH_TOKEN and TWILIO_WHATSAPP_FROM and to_number):
        return False
    url = f"https://api.twilio.com/2010-04-01/Accounts/{TWILIO_ACCOUNT_SID}/Messages.json"
    data = {
        'From': TWILIO_WHATSAPP_FROM,
        'To': to_number
    }
    if TWILIO_CONTENT_SID:
        # Gửi bằng Content Template SID
        data['ContentSid'] = TWILIO_CONTENT_SID
        if variables:
            try:
                data['ContentVariables'] = json.dumps(variables, ensure_ascii=False)
            except Exception:
                pass
    else:
        # Gửi freeform body (chỉ hoạt động trong 24h session)
        data['Body'] = body
    try:
        resp = requests.post(url, data=data, auth=(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN), timeout=10)
        return resp.status_code in (200, 201)
    except Exception:
        return False

def _get_recipients_for_row(sheet_name: str, row: dict) -> List[str]:
    """Xác định danh sách số WhatsApp cần gửi dựa vào sheet và cột đầu mối.

    - Sizing: dùng "Đầu mối xử lý" để chọn số chính, đồng thời luôn thêm số của haipn và thongnv31.
    - CapPhat: dùng "Đầu mối P.HT" để chọn số chính, đồng thời luôn thêm số của haipn và thongnv31.
    - Nếu có cấu hình trong `cache/phone_recipients.json` theo row_id, sẽ ưu tiên thêm vào danh sách.
    - Rà trùng, chỉ giữ số hợp lệ bắt đầu bằng "whatsapp:+".
    """
    recipients_cfg = _load_phone_recipients()
    res = []
    rid = row.get('row_id')
    if sheet_name == 'Sizing':
        key = str(row.get('Đầu mối xử lý', '')).strip()
    elif sheet_name == 'CapPhat':
        key = str(row.get('Đầu mối P.HT', '')).strip()
    else:
        key = ''
        # Normalize: ductn -> ductn8
        if key.lower() == 'ductn':
            key = 'ductn8'
    # Số chính theo đầu mối
    if key:
        num = OWNER_PHONE_MAP.get(key)
        if num:
            res.append(num)
    # Luôn gửi tới các đầu mối luôn nhận
    for always in ALWAYS_NOTIFY:
        num = OWNER_PHONE_MAP.get(always)
        if num:
            res.append(num)
    # Nếu có mapping theo row_id, thêm vào
    if rid:
        user_num = recipients_cfg.get(rid)
        if user_num:
            res.append(user_num)
    # Nếu không có ai, dùng mặc định môi trường
    default_to = os.environ.get('WHATSAPP_DEFAULT_TO', '')
    if not res and default_to:
        res.append(default_to)
    # Lọc hợp lệ và unique
    uniq = []
    seen = set()
    for n in res:
        n = str(n).strip()
        if not n:
            continue
        if not n.startswith('whatsapp:+'):
            continue
        if n not in seen:
            uniq.append(n)
            seen.add(n)
    return uniq

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

def _build_kpi_overdue_body(row: dict, days_overdue: int) -> str:
    proj = str(row.get('Tên dự án - Mục đích sizing', '')).strip()
    kpi = str(row.get('Thời gian hoàn thành theo KPI', '')).strip()
    label = f"Muộn {days_overdue} ngày" if days_overdue > 0 else "Quá hạn"
    return (
        f"CẢNH BÁO TIẾN ĐỘ: {label}\n"
        f"Dự án: {proj}\n"
        f"KPI: {kpi}\n"
        f"Vui lòng xử lý gấp để không ảnh hưởng tiến độ."
    )

def _build_sr_reminder_body(row: dict) -> str:
    proj = str(row.get('Dự án', '')).strip()
    received = str(row.get('Thời gian tiếp nhận y/c', '')).strip()
    return (
        f"NHẮC TẠO MÃ SR\n"
        f"Dự án: {proj}\n"
        f"Tiếp nhận: {received}\n"
        f"Yêu cầu: Vui lòng tạo mã SR trong ngày hoặc muộn nhất ngày hôm sau."
    )

def _build_sr_overdue_body(row: dict) -> str:
    proj = str(row.get('Dự án', '')).strip()
    received = str(row.get('Thời gian tiếp nhận y/c', '')).strip()
    return (
        f"ĐÃ ĐẾN HẠN TẠO MÃ SR\n"
        f"Dự án: {proj}\n"
        f"Tiếp nhận: {received}\n"
        f"Vui lòng tạo mã SR ngay để đảm bảo tiến độ."
    )

def _build_sr_deadline_body(row: dict, due_label: str, created_str: str) -> str:
    proj = str(row.get('Dự án', '')).strip()
    return (
        f"NHẮC TIẾN ĐỘ MÃ SR ({due_label})\n"
        f"Dự án: {proj}\n"
        f"Ngày tạo mã SR: {created_str}\n"
        f"YÊU CẦU: THEO DÕI TIẾN ĐỘ DỰ ÁN."
    )

def check_and_send_whatsapp_alerts():
    """Quét sheet Sizing, gửi WhatsApp cho các dòng 'Đến hạn' hoặc 'Còn 1 ngày'."""
    _refresh_sizing_progress()
    recipients = _load_phone_recipients()  
    sent = 0
    for row in data_store.get('Sizing', []):
        status = str(row.get('Tiến độ', '')).strip()
        to_numbers = _get_recipients_for_row('Sizing', row)
        if not to_numbers:
            continue
        # Gửi cảnh báo đến hạn / còn 1 ngày
        if status in ['Đến hạn', 'Còn 1 ngày']:
            for num in to_numbers:
                base_body = _build_whatsapp_body(row)
                vars0 = {
                    "project": str(row.get('Tên dự án - Mục đích sizing', '')).strip(),
                    "kpi": str(row.get('Thời gian hoàn thành theo KPI', '')).strip(),
                    "status": status
                }
                body, vars1 = _prepare_message_for_recipient('Sizing', row, num, base_body, vars0)
                if _send_whatsapp(num, body, variables=vars1):
                    sent += 1
        # Thêm cảnh báo muộn 1 ngày, muộn 2 ngày theo KPI
        kpi_dt = _parse_date(row.get('Thời gian hoàn thành theo KPI', ''))
        if kpi_dt:
            today = pd.Timestamp('today').normalize()
            kpi_norm = pd.Timestamp(kpi_dt).normalize()
            overdue_days = (today - kpi_norm).days
            if overdue_days in [1, 2]:
                for num in to_numbers:
                    base_body = _build_kpi_overdue_body(row, overdue_days)
                    vars0 = {
                        "project": str(row.get('Tên dự án - Mục đích sizing', '')).strip(),
                        "kpi": str(row.get('Thời gian hoàn thành theo KPI', '')).strip(),
                        "status": f"Muộn {overdue_days} ngày"
                    }
                    body, vars1 = _prepare_message_for_recipient('Sizing', row, num, base_body, vars0)
                    if _send_whatsapp(num, body, variables=vars1):
                        sent += 1
    # Quét sheet Cấp phát TN: nếu đã tiếp nhận hôm nay hoặc hôm qua mà chưa có 'Mã SR' -> nhắc tạo SR
    try:
        today = pd.Timestamp('today').normalize()
        for row in data_store.get('CapPhat', []):
            sr = str(row.get('Mã SR', '')).strip()
            recv = _parse_date(row.get('Thời gian tiếp nhận y/c', ''))
            if not sr and recv:
                recv = pd.Timestamp(recv).normalize()
                delta_days = (today - recv).days
                if 0 <= delta_days <= 3:
                    to_numbers = _get_recipients_for_row('CapPhat', row)
                    if to_numbers:
                        if delta_days == 0:
                            # Ngày tiếp nhận: gửi nhắc tạo SR
                            for num in to_numbers:
                                base_body = _build_sr_reminder_body(row)
                                vars0 = {
                                    "project": str(row.get('Dự án', '')).strip(),
                                    "received": str(row.get('Thời gian tiếp nhận y/c', '')).strip()
                                }
                                body, vars1 = _prepare_message_for_recipient('CapPhat', row, num, base_body, vars0)
                                if _send_whatsapp(num, body, variables=vars1):
                                    sent += 1
                        elif delta_days == 1:
                            # Ngày hôm sau: gửi cảnh báo đã đến hạn tạo SR
                            for num in to_numbers:
                                base_body = _build_sr_overdue_body(row)
                                vars0 = {
                                    "project": str(row.get('Dự án', '')).strip(),
                                    "received": str(row.get('Thời gian tiếp nhận y/c', '')).strip()
                                }
                                body, vars1 = _prepare_message_for_recipient('CapPhat', row, num, base_body, vars0)
                                if _send_whatsapp(num, body, variables=vars1):
                                    sent += 1
                        elif delta_days in [2, 3]:
                            # Muộn 1-2 ngày kể từ hạn tạo SR nếu chưa tạo mã SR
                            label = f"Muộn {delta_days-1} ngày"
                            body = (
                                f"NHẮC TẠO MÃ SR ({label})\n"
                                f"Dự án: {str(row.get('Dự án', '')).strip()}\n"
                                f"Tiếp nhận: {str(row.get('Thời gian tiếp nhận y/c', '')).strip()}\n"
                                f"Vui lòng tạo mã SR ngay để đảm bảo tiến độ."
                            )
                            for num in to_numbers:
                                vars0 = {
                                    "project": str(row.get('Dự án', '')).strip(),
                                    "received": str(row.get('Thời gian tiếp nhận y/c', '')).strip(),
                                    "status": label
                                }
                                body2, vars1 = _prepare_message_for_recipient('CapPhat', row, num, body, vars0)
                                if _send_whatsapp(num, body2, variables=vars1):
                                    sent += 1
    except Exception:
        pass
    # Nhắc tiến độ sau khi đã có Mã SR: Deadline = 2 ngày sau ngày tạo mã SR (calendar days)
    try:
        sr_map = _load_sr_created_map()
        today = pd.Timestamp('today').normalize()
        for row in data_store.get('CapPhat', []):
            sr = str(row.get('Mã SR', '')).strip()
            rid = row.get('row_id')
            if sr and rid and rid in sr_map:
                created_str = sr_map.get(rid, '')
                created_dt = _parse_date(created_str)
                if not created_dt:
                    continue
                created_dt = pd.Timestamp(created_dt).normalize()
                deadline = created_dt + pd.Timedelta(days=2)
                days_left = (deadline - today).days
                if days_left in [0, 1]:
                    due_label = 'Đến hạn' if days_left == 0 else 'Còn 1 ngày'
                    to_numbers = _get_recipients_for_row('CapPhat', row)
                    for num in to_numbers:
                        base_body = _build_sr_deadline_body(row, due_label, created_str)
                        vars0 = {
                            "project": str(row.get('Dự án', '')).strip(),
                            "created": created_str,
                            "due": due_label
                        }
                        body, vars1 = _prepare_message_for_recipient('CapPhat', row, num, base_body, vars0)
                        if _send_whatsapp(num, body, variables=vars1):
                            sent += 1
    except Exception:
        pass
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

# --- Cấu hình sheet "Tài nguyên Cloud" ---
CLOUD_COLUMNS = [
    # Khối QHĐC 2020
    "STT","Tiêu chí",
    "Com_Ceph vCPU","Com_SAN vCPU",
    "Com_Ceph RAM(TB)","Com_SAN RAM(TB)",
    "CEPH(GB)","SAN(GB)","NAS(GB)","Ghi chú",

    # Khối Private Cloud
    "PC Tiêu chí",
    "PC Com_Ceph vCPU","PC Com_SAN vCPU",
    "PC Com_Ceph RAM(GB)","PC Com_SAN RAM(GB)",
    "PC CEPH(GB)","PC SAN(GB)","PC NAS(GB)","NAS(theo đầu mới VHKTT gửi tháng 3 2025)","PC Ghi chú",

    # Số liệu tổng hợp năm
    "Tổng TN 2021","Đã cấp phát 2021","TN còn lại 2021",
    "Tổng TN 2023","Đã cấp phát 2023","TN còn lại 2023",

    # Khối QHĐC 2021
    "2021 Tiêu chí",
    "2021 Com_Ceph vCPU","2021 Com_SAN vCPU",
    "2021 Com_Ceph RAM(GB)","2021 Com_SAN RAM(GB)",
    "2021 CEPH(GB)","2021 SAN(GB)","2021 NAS(GB)","2021 NAS(theo đầu mới VHKTT gửi tháng 3 2025)","2021 Ghi chú",

    # Khối QHĐC 2022
    "2022 vCPU","2022 RAM(GB)","2022 CEPH(GB)","2022 Com_SAN","2022 SAN(GB)","2022 NAS(GB)","Object(GB)",
    "Com_Bigdata vCPU","Com_Bigdata RAM","Bigdata(GB)","Archiving(GB)","Bare_metal vCPU","Bare_metal RAM","2022 Ghi chú"
]

"""Bộ nhớ dữ liệu chính trong runtime (3 sheet)."""
data_store = {
    'Sizing': initial_rows(SIZING_COLUMNS),
    'CapPhat': initial_rows(CAP_PHAT_COLUMNS),
    'ChiTiet': initial_rows(CHI_TIET_COLUMNS),
    'Cloud': initial_rows(CLOUD_COLUMNS)
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
ensure_stt(data_store['Cloud'])

def ensure_cloud_min_rows():
    """Đảm bảo sheet Cloud có tối thiểu 6 hàng để đủ nhãn (2020 + Private Cloud)."""
    target = data_store.get('Cloud', [])
    needed = 6 - len(target)
    if needed > 0:
        for _ in range(needed):
            target.append(blank_row(CLOUD_COLUMNS))
        ensure_stt(target)
    data_store['Cloud'] = target

ensure_cloud_min_rows()

"""Thiết lập đường dẫn cache để lưu/khôi phục `data_store`."""
CACHE_DIR = os.path.join(os.path.dirname(__file__), 'cache')
os.makedirs(CACHE_DIR, exist_ok=True)
CACHE_FILE = os.path.join(CACHE_DIR, 'data_store.json')
PHONE_RECIPIENTS_FILE = os.path.join(CACHE_DIR, 'phone_recipients.json')  # mapping row_id -> whatsapp phone
CAP_PHAT_SR_CREATED_FILE = os.path.join(CACHE_DIR, 'cap_phat_sr_created.json')  # mapping row_id -> Ngày tạo mã SR (dd/mm/YYYY)

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

def _load_sr_created_map():
    if os.path.exists(CAP_PHAT_SR_CREATED_FILE):
        try:
            with open(CAP_PHAT_SR_CREATED_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def _save_sr_created_map(mapping: dict):
    try:
        with open(CAP_PHAT_SR_CREATED_FILE, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
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
            if 'Cloud' in loaded:
                data_store['Cloud'] = sanitize_rows(loaded['Cloud'], CLOUD_COLUMNS)
                ensure_stt(data_store['Cloud'])
                ensure_cloud_min_rows()
        except Exception:
            pass

load_cache()

"""Trang chính: render giao diện với 3 bảng dữ liệu."""
@app.route('/')
def index():
    _refresh_sizing_progress()
    # Tính bảng tổng hợp Sizing theo quý/năm và Owner
    def _quarter_of(dt: pd.Timestamp) -> int:
        m = int(dt.month)
        return 1 if m<=3 else (2 if m<=6 else (3 if m<=9 else 4))
    allowed_owners = ["khanhnd23","ductn8","vinhtq18","thongnv31","tuanha3"]
    owners = []
    year_set = set()
    counts = {}
    reject_counts = {}
    for r in data_store.get('Sizing', []):
        owner = str(r.get('Đầu mối xử lý','')).strip() or ''
        if owner and owner in allowed_owners:
            if owner not in owners:
                owners.append(owner)
        # đếm trạng thái Từ chối
        status = str(r.get('Trạng thái','')).strip().lower()
        if owner and owner in allowed_owners:
            reject_counts[owner] = reject_counts.get(owner, 0) + (1 if status == 'từ chối' else 0)
        # quý/năm từ Thời điểm đẩy yêu cầu
        dt = _parse_date(r.get('Thời điểm đẩy yêu cầu',''))
        if dt:
            y = int(pd.Timestamp(dt).year)
            # Bỏ qua các năm không hợp lệ (ví dụ: 1970 do lỗi parse Excel)
            if y < 2000 or y > 2100:
                continue
            q = _quarter_of(pd.Timestamp(dt))
            year_set.add(y)
            key = (owner, y, q)
            if owner in allowed_owners:
                counts[key] = counts.get(key, 0) + 1
    years = sorted(list(year_set))
    # Chuẩn hoá dữ liệu cho template
    sizing_summary_rows = []
    # Cố định thứ tự theo allowed_owners
    owners = [o for o in allowed_owners if o in owners]
    for owner in owners:
        row = {'owner': owner, 'quarters': {}, 'owner_false': reject_counts.get(owner, 0)}
        for y in years:
            row['quarters'][y] = {
                1: counts.get((owner, y, 1), 0),
                2: counts.get((owner, y, 2), 0),
                3: counts.get((owner, y, 3), 0),
                4: counts.get((owner, y, 4), 0)
            }
        sizing_summary_rows.append(row)
    # Hàng tổng
    total_row = {'owner': 'Tổng', 'quarters': {}, 'owner_false': sum(reject_counts.values())}
    for y in years:
        total_row['quarters'][y] = {
            1: sum(counts.get((o, y, 1), 0) for o in owners),
            2: sum(counts.get((o, y, 2), 0) for o in owners),
            3: sum(counts.get((o, y, 3), 0) for o in owners),
            4: sum(counts.get((o, y, 4), 0) for o in owners)
        }
    sizing_summary_rows.append(total_row)
    # Tính tổng hợp Chi tiết theo 'Nhóm tài nguyên'
    CHI_TIET_SUM_COLS = [
        'vCPU','Cint','RAM(GB)','SAN(GB)','NAS(GB)','Ceph(GB)','Bigdata(GB)','Archiving(GB)','S3 Object(GB)'
    ]
    chitiet_group_totals = {}
    for r in data_store.get('ChiTiet', []):
        group = str(r.get('Nhóm tài nguyên','')).strip() or ''
        if not group:
            continue
        if group not in chitiet_group_totals:
            chitiet_group_totals[group] = {c: 0 for c in CHI_TIET_SUM_COLS}
        for c in CHI_TIET_SUM_COLS:
            try:
                val = r.get(c, '')
                num = pd.to_numeric(val, errors='coerce')
                if not pd.isna(num):
                    chitiet_group_totals[group][c] += float(num)
            except Exception:
                pass
    # Hàng tổng cộng
    chitiet_total_row = {c: 0 for c in CHI_TIET_SUM_COLS}
    for grp, sums in chitiet_group_totals.items():
        for c in CHI_TIET_SUM_COLS:
            chitiet_total_row[c] += sums.get(c, 0)
    return render_template(
        'index.html',
        sizing_columns=SIZING_COLUMNS,
        cap_phat_columns=CAP_PHAT_COLUMNS,
        chi_tiet_columns=CHI_TIET_COLUMNS,
        cloud_columns=CLOUD_COLUMNS,
        sizing_rows=data_store['Sizing'],
        cap_phat_rows=data_store['CapPhat'],
        chi_tiet_rows=data_store['ChiTiet'],
        cloud_rows=data_store['Cloud'],
        sizing_summary_years=years,
        sizing_summary_rows=sizing_summary_rows,
        chitiet_group_totals=chitiet_group_totals,
        chitiet_sum_cols=CHI_TIET_SUM_COLS,
        chitiet_total_row=chitiet_total_row
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

    # Helper làm sạch chuỗi và chuẩn hoá giá trị theo cột
    def _normalize_cell(col_name: str, value: str) -> str:
        s = str(value)
        s = s.strip()
        if not s:
            return ""
        # Chuẩn hoá owner ductn -> ductn8 cho sheet Sizing
        if col_name == 'Đầu mối xử lý' and s.lower() == 'ductn':
            return 'ductn8'
        return s

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
            # Không cố parse ngày ở các cột còn lại; giữ nguyên như chuỗi sạch và chuẩn hoá cần thiết.
            processed[col] = series.astype(str).apply(
                lambda x: '' if str(x).strip().lower() in ['nan','nat'] else _normalize_cell(col, x)
            )

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
        cloud_df = pd.read_excel(xl, sheet_name='Tài nguyên Cloud') if 'Tài nguyên Cloud' in xl.sheet_names else pd.DataFrame(columns=CLOUD_COLUMNS)

        sizing_df = _read_sheet(sizing_df, SIZING_COLUMNS)
        cap_phat_df = _read_sheet(cap_phat_df, CAP_PHAT_COLUMNS)
        chi_tiet_df = _read_sheet(chi_tiet_df, CHI_TIET_COLUMNS)
        cloud_df = _read_sheet(cloud_df, CLOUD_COLUMNS)

        data_store['Sizing'] = sanitize_rows(sizing_df.to_dict(orient='records'), SIZING_COLUMNS)
        data_store['CapPhat'] = sanitize_rows(cap_phat_df.to_dict(orient='records'), CAP_PHAT_COLUMNS)
        data_store['ChiTiet'] = sanitize_rows(chi_tiet_df.to_dict(orient='records'), CHI_TIET_COLUMNS)
        data_store['Cloud'] = sanitize_rows(cloud_df.to_dict(orient='records'), CLOUD_COLUMNS)
        _fix_chitiet_numeric_rows(data_store['ChiTiet'])

        _refresh_sizing_progress()

        ensure_stt(data_store['Sizing'])
        ensure_stt(data_store['CapPhat'])
        ensure_stt(data_store['ChiTiet'])
        ensure_stt(data_store['Cloud'])
        ensure_cloud_min_rows()
        save_cache()

        return render_template('tables.html', sizing_columns=SIZING_COLUMNS, cap_phat_columns=CAP_PHAT_COLUMNS,
                       chi_tiet_columns=CHI_TIET_COLUMNS, cloud_columns=CLOUD_COLUMNS,
                       sizing_rows=data_store['Sizing'], cap_phat_rows=data_store['CapPhat'], chi_tiet_rows=data_store['ChiTiet'], cloud_rows=data_store['Cloud'])
    except Exception as e:
        return jsonify({'error': str(e)}), 500

"""Render partial cho một sheet bất kỳ (HTMX sử dụng)."""
@app.route('/sheet/<name>')
def sheet(name):
    if name not in ['Sizing', 'CapPhat', 'ChiTiet', 'Cloud']:
        abort(404)
    if name == 'Sizing':
        columns = SIZING_COLUMNS
        sheet_id = 'sizing-sheet'
    elif name == 'CapPhat':
        columns = CAP_PHAT_COLUMNS
        sheet_id = 'cap-phat-sheet'
    elif name == 'ChiTiet':
        columns = CHI_TIET_COLUMNS
        sheet_id = 'chi-tiet-sheet'
    else:
        columns = CLOUD_COLUMNS
        sheet_id = 'cloud-sheet'
    if name == 'Sizing':
        _refresh_sizing_progress()
    rows = data_store[name]
    return render_template('sheet.html', sheet_name=name, sheet_id=sheet_id, columns=columns, rows=rows)

"""Thêm một hàng mới sau vị trí chỉ định trong sheet."""
@app.route('/add-row/<sheet>/<int:after_index>', methods=['POST'])
def add_row(sheet, after_index):
    if sheet not in ['Sizing', 'CapPhat', 'ChiTiet', 'Cloud']:
        abort(404)
    columns = SIZING_COLUMNS if sheet == 'Sizing' else (CAP_PHAT_COLUMNS if sheet == 'CapPhat' else (CHI_TIET_COLUMNS if sheet == 'ChiTiet' else CLOUD_COLUMNS))
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
    sheet_id = 'sizing-sheet' if sheet == 'Sizing' else ('cap-phat-sheet' if sheet == 'CapPhat' else ('chi-tiet-sheet' if sheet == 'ChiTiet' else 'cloud-sheet'))
    return render_template('sheet.html', sheet_name=sheet, sheet_id=sheet_id, columns=columns, rows=target_list)

"""Xoá một hàng theo chỉ số trong sheet."""
@app.route('/delete-row/<sheet>/<int:row_index>', methods=['POST'])
def delete_row(sheet, row_index):
    if sheet not in ['Sizing', 'CapPhat', 'ChiTiet', 'Cloud']:
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
    columns = SIZING_COLUMNS if sheet == 'Sizing' else (CAP_PHAT_COLUMNS if sheet == 'CapPhat' else (CHI_TIET_COLUMNS if sheet == 'ChiTiet' else CLOUD_COLUMNS))
    sheet_id = 'sizing-sheet' if sheet == 'Sizing' else ('cap-phat-sheet' if sheet == 'CapPhat' else ('chi-tiet-sheet' if sheet == 'ChiTiet' else 'cloud-sheet'))
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

    if sheet not in ['Sizing', 'CapPhat', 'ChiTiet', 'Cloud']:
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
    columns = SIZING_COLUMNS if sheet == 'Sizing' else (CAP_PHAT_COLUMNS if sheet == 'CapPhat' else (CHI_TIET_COLUMNS if sheet == 'ChiTiet' else CLOUD_COLUMNS))
    if col not in columns:
        return jsonify({'error': 'Invalid column'}), 400
    prev_val = target_row.get(col, '')
    target_row[col] = value
    if sheet == 'Sizing' and col == 'Thời gian hoàn thành theo KPI':
        _update_progress_for_row(target_row)
    # Khi tạo Mã SR (CapPhat) từ rỗng -> có giá trị: ghi lại 'Ngày tạo mã SR' (ẩn) vào map
    if sheet == 'CapPhat' and col == 'Mã SR':
        try:
            rid = target_row.get('row_id')
            if rid and (not str(prev_val).strip()) and str(value).strip():
                sr_map = _load_sr_created_map()
                sr_map[rid] = datetime.now().strftime('%d/%m/%Y')
                _save_sr_created_map(sr_map)
        except Exception:
            pass
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
            cloud_df = build_df(data_store['Cloud'], CLOUD_COLUMNS)

            sizing_df.to_excel(writer, sheet_name='Sizing', index=False)
            cap_df.to_excel(writer, sheet_name='Cấp phát TN', index=False)
            chi_tiet_df.to_excel(writer, sheet_name='Chi tiết', index=False)
            cloud_df.to_excel(writer, sheet_name='Tài nguyên Cloud', index=False)
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
    elif sheet_name == 'Cloud':
        return data_store['Cloud'], CLOUD_COLUMNS, 'cloud-sheet'
    abort(404)

"""Cập nhật danh sách cột chuẩn tương ứng theo tên sheet."""
def update_columns_constant(sheet_name, new_columns):
    global SIZING_COLUMNS, CAP_PHAT_COLUMNS, CHI_TIET_COLUMNS, CLOUD_COLUMNS

    if sheet_name == 'Sizing':
        SIZING_COLUMNS = new_columns
    elif sheet_name == 'CapPhat':
        CAP_PHAT_COLUMNS = new_columns
    elif sheet_name == 'ChiTiet':
        CHI_TIET_COLUMNS = new_columns
    elif sheet_name == 'Cloud':
        CLOUD_COLUMNS = new_columns

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
    import threading, time

    FIXED_TIMES = [
        (9, 0),   
        (14, 0),
        (16, 30), 
    ]

    def next_schedule_after(now: datetime) -> datetime:
        today = now.date()
        # Tìm mốc thời gian gần nhất còn lại trong ngày
        for h, m in FIXED_TIMES:
            candidate = datetime(year=today.year, month=today.month, day=today.day, hour=h, minute=m, second=0, microsecond=0)
            if candidate > now:
                return candidate
        # Nếu không còn mốc trong ngày, chuyển sang mốc đầu tiên của ngày hôm sau
        tomorrow = now + timedelta(days=1)
        ymd = tomorrow.date()
        h, m = FIXED_TIMES[0]
        return datetime(year=ymd.year, month=ymd.month, day=ymd.day, hour=h, minute=m, second=0, microsecond=0)

    def loop():
        # Không gửi ngay khi khởi động; chỉ gửi theo các mốc cố định.
        while True:
            try:
                now = datetime.now()
                next_run = next_schedule_after(now)
                sleep_seconds = (next_run - now).total_seconds()
                time.sleep(max(sleep_seconds, 0))
                # Đến mốc giờ: gửi cảnh báo
                try:
                    check_and_send_whatsapp_alerts()
                except Exception:
                    pass
            except Exception:
                # Nếu có lỗi bất ngờ, ngủ 1 phút rồi thử lại để không chết thread
                time.sleep(60)

    t = threading.Thread(target=loop, name='whatsapp-scheduler', daemon=True)
    t.start()

"""Endpoint debug: xem cấu hình môi trường hiện tại."""
@app.route('/debug-env', methods=['GET'])
def debug_env():
    cfg = {
        'TWILIO_ACCOUNT_SID': os.environ.get('TWILIO_ACCOUNT_SID', ''),
        'TWILIO_AUTH_TOKEN': os.environ.get('TWILIO_AUTH_TOKEN', ''),
        'TWILIO_WHATSAPP_FROM': os.environ.get('TWILIO_WHATSAPP_FROM', ''),
        'TWILIO_CONTENT_SID': os.environ.get('TWILIO_CONTENT_SID', ''),
        'WHATSAPP_DEFAULT_TO': os.environ.get('WHATSAPP_DEFAULT_TO', ''),
    }
    try:
        print('[DEBUG-ENV]', cfg)
    except Exception:
        pass
    return jsonify(cfg), 200

"""Điểm vào ứng dụng (chạy development server)."""
if __name__ == '__main__':
    # Khởi động scheduler tự động gửi WhatsApp nếu đã cấu hình Twilio FROM
    if TWILIO_WHATSAPP_FROM:
        _start_whatsapp_daily_scheduler()
    app.run(debug=True)