import streamlit as st
import streamlit.components.v1 as components
import datetime
import io
import requests
import hashlib
import hmac
import secrets
from docxtpl import DocxTemplate
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt


# ==========================================
# ⚠️ ตั้งค่า URL ที่ต้องใช้งาน 2 จุดที่นี่ค่ะ ⚠️
# ==========================================
# 1. URL จาก Google Apps Script (เว็บแอป)
WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbwnevwKC_8BWNNKPIvzQ_F2AngKpFambc4U0SFTRjo4ZY0z6tJXhwyugmwd-pMz8Tdh/exec"

# 2. URL จาก Looker Studio (ฝังรายงาน)
DASHBOARD_URL = "https://datastudio.google.com/embed/reporting/eb15729f-a596-4adc-bb19-b28269e6838d/page/g566F"
# ==========================================

# --- ฟังก์ชันช่วยเหลือ ---
def get_thai_date(target_date):
    thai_months = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
    return f"{target_date.day} {thai_months[target_date.month]} {target_date.year + 543}"


# --- ฟังก์ชันช่วยจัดการหลายรายการยาในเอกสารเดียว ---
UNITS = ["เม็ด", "ไวอัล", "แอมพูล", "ขวด", "หลอด", "กล่อง", "แกลลอน", "set", "ชิ้น", "อื่นๆ"]

def get_transaction_no(state_key, prefix):
    """สร้างเลขเอกสารครั้งเดียวต่อ 1 transaction เพื่อไม่ให้เลขเปลี่ยนเมื่อ Streamlit rerun"""
    if state_key not in st.session_state:
        st.session_state[state_key] = f"{prefix}-{datetime.datetime.now().strftime('%y%m%d-%H%M%S')}"
    return st.session_state[state_key]

def get_item_list(state_key):
    if state_key not in st.session_state:
        st.session_state[state_key] = []
    return st.session_state[state_key]

def format_qty(value):
    """แสดงจำนวนเต็มโดยไม่ติด .0 แต่ยังรองรับจำนวนทศนิยม"""
    try:
        value = float(value)
        return str(int(value)) if value.is_integer() else f"{value:g}"
    except (TypeError, ValueError):
        return str(value)

def make_line_doc_no(base_doc_no, index):
    return f"{base_doc_no}-{index:02d}"

def drug_details_text(items, opd_mode=False):
    lines = []
    for idx, item in enumerate(items, start=1):
        if opd_mode:
            lines.append(
                f"{idx}. {item['drug_name']} - จ่ายให้ผู้ป่วย 3 วัน {format_qty(item['qty_3day'])} {item['unit']} "
                f"/ รพ.ปลายทางทำเรื่องยืม {format_qty(item['borrow_qty'])} {item['unit']}"
            )
        else:
            lines.append(
                f"{idx}. {item['drug_name']} จำนวน {format_qty(item['qty'])} {item['unit']}"
            )
    return "\n".join(lines)


def _set_cell_text(cell, text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT):
    """ใส่ข้อความใน cell ด้วย TH Sarabun New ให้ตรงกับ template งานสารบรรณ"""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    run = p.add_run(str(text))
    run.bold = bold
    run.font.name = "TH Sarabun New"
    run.font.size = Pt(16)

    # Word แยกการกำหนด font สำหรับ Latin / Thai / complex script
    # จึงระบุให้ครบเพื่อป้องกันข้อความที่สร้างใหม่เด้งกลับไปเป็น Calibri/Aptos
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), "TH Sarabun New")
    rfonts.set(qn("w:hAnsi"), "TH Sarabun New")
    rfonts.set(qn("w:eastAsia"), "TH Sarabun New")
    rfonts.set(qn("w:cs"), "TH Sarabun New")


def _hide_table_borders(table):
    """ซ่อนเส้นตารางทั้งหมด แต่ยังคงโครงสร้างคอลัมน์ไว้"""
    tblPr = table._tbl.tblPr
    borders = tblPr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "nil")


def replace_drug_marker_with_table(docx_bytes, items, opd_mode=False, marker="__DRUG_TABLE__"):
    """แทน marker ใน Word ด้วยตารางไร้เส้นสำหรับรายการยา"""
    docx_bytes.seek(0)
    doc = Document(docx_bytes)

    marker_paragraph = None
    for paragraph in doc.paragraphs:
        if marker in paragraph.text:
            marker_paragraph = paragraph
            break

    # หาก template ไม่มี marker ให้คงเอกสารเดิมไว้เพื่อไม่ให้ระบบล้ม
    if marker_paragraph is None:
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        return output

    # ลบเฉพาะ marker แต่คงข้อความอื่นใน paragraph เดิม (ถ้ามี)
    for run in marker_paragraph.runs:
        if marker in run.text:
            run.text = run.text.replace(marker, "")

    if opd_mode:
        # คอลัมน์สุดท้ายช่วยตรวจสอบได้ทันทีว่า 3 วันที่ รพ.เราจ่าย + ส่วนที่ รพ.ปลายทางยืม
        # รวมแล้วตรงกับจำนวนยาที่ผู้ป่วยต้องใช้ทั้งหมดหรือไม่
        headers = ["ลำดับ", "รายการยา", "จ่ายผู้ป่วย\n3 วัน", "รพ.ปลายทาง\nขอยืม", "รวมที่ต้องใช้"]
        rows = []
        for idx, item in enumerate(items, start=1):
            qty_3day = float(item["qty_3day"])
            borrow_qty = float(item["borrow_qty"])
            total_qty = qty_3day + borrow_qty
            rows.append([
                idx,
                item["drug_name"],
                f"{format_qty(qty_3day)} {item['unit']}",
                f"{format_qty(borrow_qty)} {item['unit']}",
                f"{format_qty(total_qty)} {item['unit']}",
            ])
    else:
        headers = ["ลำดับ", "รายการยา", "จำนวน"]
        rows = [
            [idx, item["drug_name"], f"{format_qty(item['qty'])} {item['unit']}"]
            for idx, item in enumerate(items, start=1)
        ]

    table = doc.add_table(rows=1, cols=len(headers))
    table.autofit = True
    _hide_table_borders(table)

    for col, heading in enumerate(headers):
        _set_cell_text(table.rows[0].cells[col], heading, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)

    for row_values in rows:
        cells = table.add_row().cells
        for col, value in enumerate(row_values):
            align = WD_ALIGN_PARAGRAPH.CENTER if col == 0 else WD_ALIGN_PARAGRAPH.LEFT
            _set_cell_text(cells[col], value, align=align)

    # ย้ายตารางจากท้ายเอกสารไปไว้ตรงตำแหน่ง {{ drug_details }} เดิม
    # แล้วลบ paragraph ของ marker ทิ้งทั้งหมด เพื่อไม่ให้เลขลำดับ/list เช่น "1."
    # ที่ติดมาจาก template ค้างอยู่เหนือรายการยา
    marker_p = marker_paragraph._p
    marker_p.addnext(table._tbl)
    marker_p.getparent().remove(marker_p)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- ฟังก์ชันติดต่อ Google Sheets ---
def get_from_google_sheets(sheet_name):
    try:
        response = requests.get(
            WEBHOOK_URL,
            params={"sheet_name": sheet_name},
            timeout=20,
        )
        response.raise_for_status()
        result = response.json()
        if result.get("status") == "success":
            return result.get("data", [])
        # ไม่ให้ error ถูกกลืนจนหา原因ไม่ได้
        st.session_state[f"sheet_error_{sheet_name}"] = result.get("message", "Unknown Google Sheets error")
        return []
    except Exception as e:
        st.session_state[f"sheet_error_{sheet_name}"] = str(e)
        return []

def save_to_google_sheets(sheet_name, row_data=None, action="append", doc_id=None, new_status=None):
    if "script.google.com" not in WEBHOOK_URL:
        return False, "ยังไม่ได้ใส่ WEBHOOK_URL ในโค้ด Python ค่ะ"

    payload = {"sheet_name": sheet_name, "action": action}
    if action == "append":
        payload["row_data"] = row_data
    elif action == "update":
        payload["doc_id"] = doc_id
        payload["new_status"] = new_status

    try:
        response = requests.post(
            WEBHOOK_URL,
            json=payload,
            allow_redirects=True,
            timeout=20,
        )
        response.raise_for_status()
        result = response.json()
        if result.get("status") == "success":
            return True, result.get("message", "Success")
        return False, result.get("message", response.text)
    except Exception as e:
        return False, str(e)


# ==========================================
# 🔐 ระบบรหัสผ่านรายบุคคลสำหรับเจ้าหน้าที่คลังยา
# ==========================================
# ต้องมีชีตชื่อ Admin_Passwords ใน Google Sheets เดียวกับฐานข้อมูล
# แนะนำหัวคอลัมน์แถวแรก: user_name | salt | password_hash | created_at
# ระบบจะเก็บเฉพาะ salt + hash ไม่เก็บรหัสผ่านจริง
ADMIN_PASSWORD_SHEET = "Admin_Passwords"
PASSWORD_ITERATIONS = 310_000

def _hash_password(password, salt_hex):
    salt = bytes.fromhex(salt_hex)
    return hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        PASSWORD_ITERATIONS,
    ).hex()

PASSWORD_RESET_MARKER = "__RESET__"
SYSTEM_ADMIN_SECRET_KEY = "SYSTEM_ADMIN_PASSWORD"


def get_latest_admin_credential(user_name):
    """อ่าน credential ล่าสุดของเจ้าหน้าที่คลังจาก Google Sheets

    ถ้า record ล่าสุดเป็น RESET marker จะถือว่าบัญชีนั้นยังไม่มีรหัสผ่าน
    และให้เจ้าของบัญชีตั้งรหัสใหม่ในการเข้าใช้งานครั้งถัดไป
    """
    rows = get_from_google_sheets(ADMIN_PASSWORD_SHEET)
    if not rows:
        return None

    # ใช้ record ล่าสุดของผู้ใช้นั้นเท่านั้น เพื่อไม่ย้อนกลับไปใช้รหัสเก่าหลังถูก reset
    for row in reversed(rows):
        if len(row) >= 1 and str(row[0]).strip() == user_name:
            salt_hex = str(row[1]).strip() if len(row) > 1 else ""
            password_hash = str(row[2]).strip() if len(row) > 2 else ""

            if salt_hex == PASSWORD_RESET_MARKER or password_hash == PASSWORD_RESET_MARKER:
                return None

            if salt_hex and password_hash:
                return {"salt": salt_hex, "password_hash": password_hash}

            # หาก record ล่าสุดของผู้ใช้นี้ไม่สมบูรณ์ ให้ถือว่าไม่มี credential
            # เพื่อไม่เผลอย้อนกลับไปยอมรับรหัสเก่า
            return None
    return None

def verify_admin_password(user_name, password):
    credential = get_latest_admin_credential(user_name)
    if not credential:
        return False
    try:
        candidate_hash = _hash_password(password, credential["salt"])
        return hmac.compare_digest(candidate_hash, credential["password_hash"])
    except (ValueError, TypeError):
        return False

def set_admin_password(user_name, new_password):
    """บันทึกรหัสใหม่แบบ hash ลง Google Sheets โดยไม่ส่งรหัสจริงไปเก็บ"""
    salt_hex = secrets.token_hex(16)
    password_hash = _hash_password(new_password, salt_hex)
    timestamp = datetime.datetime.now().isoformat(timespec="seconds")
    return save_to_google_sheets(
        ADMIN_PASSWORD_SHEET,
        row_data=[user_name, salt_hex, password_hash, timestamp],
        action="append",
    )

def password_is_valid(password):
    # ตั้งขั้นต่ำไม่ยาวเกินไปเพื่อให้ใช้งานจริงสะดวก แต่ไม่อนุญาตรหัสสั้นมาก
    return len(password) >= 6


def get_system_admin_password():
    """อ่านรหัส System Admin จาก Streamlit Secrets โดยไม่เขียนรหัสไว้ใน source code"""
    try:
        return str(st.secrets[SYSTEM_ADMIN_SECRET_KEY])
    except Exception:
        return ""


def verify_system_admin_password(password):
    configured_password = get_system_admin_password()
    if not configured_password:
        return False
    return hmac.compare_digest(str(password), configured_password)


def reset_warehouse_password(user_name):
    """รีเซ็ตรหัสเจ้าหน้าที่คลังโดยบันทึก reset marker แทนการลบประวัติเดิม"""
    timestamp = datetime.datetime.now().isoformat(timespec="seconds")
    return save_to_google_sheets(
        ADMIN_PASSWORD_SHEET,
        row_data=[user_name, PASSWORD_RESET_MARKER, PASSWORD_RESET_MARKER, timestamp],
        action="append",
    )


# ==========================================
# 🖥️ เริ่มต้นการสร้างหน้าจอแอปพลิเคชัน
# ==========================================
st.set_page_config(page_title="ระบบจัดการยืม-คืนยา โรงพยาบาลศรีสังวรสุโขทัย", layout="centered", page_icon="SSW_Logo.jpg")

# --- โค้ดฝังฟอนต์ Sarabun, ซ่อนลายน้ำ และปรับแต่งสำหรับมือถือ ---
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');
    
    /* ใช้ Sarabun กับข้อความทั่วไป แต่ไม่บังคับกับ <span> ทุกตัว
       เพราะ Streamlit ใช้ Material Symbols ผ่าน span สำหรับไอคอน
       ถ้าบังคับ font-family กับ span จะเห็นคำว่า arrow_right / keyboard_double... ซ้อนกับข้อความ */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stSidebar"],
    h1, h2, h3, h4, h5, h6, p, label, input, textarea, button, select {
        font-family: 'Sarabun', sans-serif !important;
    }

    /* คืนฟอนต์ไอคอนของ Streamlit ให้ Material Symbols */
    .material-symbols-rounded {
        font-family: 'Material Symbols Rounded' !important;
        font-weight: normal !important;
        font-style: normal !important;
    }
    .material-symbols-outlined {
        font-family: 'Material Symbols Outlined' !important;
        font-weight: normal !important;
        font-style: normal !important;
    }
    [data-testid="stIconMaterial"] {
        font-family: 'Material Symbols Rounded' !important;
        font-weight: normal !important;
        font-style: normal !important;
    }
    
    #MainMenu {visibility: hidden;} 
    footer {visibility: hidden;} 
    header {visibility: hidden;}

    [data-testid="stSidebar"] img {
        max-width: 120px !important;
        display: block;
        margin-left: auto;
        margin-right: auto;
    }

    [data-testid="stSidebar"] [data-testid="stImage"] {
        display: flex;
        justify-content: center;
        align-items: center;
    }

    @media (max-width: 768px) {
        h1 { font-size: 26px !important; }
        h2 { font-size: 22px !important; }
        h3 { font-size: 18px !important; }
        
        .stButton>button {
            width: 100% !important;
            padding: 15px !important;
            font-size: 18px !important;
            border-radius: 10px !important;
        }
        
        .block-container {
            padding-left: 1rem !important;
            padding-right: 1rem !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --- หัวข้อหลักพร้อมโลโก้โรงพยาบาล ---
col_logo, col_title = st.columns([1, 6])
with col_logo:
    try:
        st.image("SSW_Logo.jpg", width=70)
    except:
        pass
with col_title:
    st.markdown(
        """
        <h3 style='margin: 0; padding-top: 2px;'>ระบบจัดการยืม-คืนยา</h3>
        <p style='margin: 0; color: #555; font-size: 18px;'>กลุ่มงานเภสัชกรรม โรงพยาบาลศรีสังวรสุโขทัย</p>
        """, 
        unsafe_allow_html=True
    )

st.markdown("---")

# 1. ฐานข้อมูลเภสัชกร 
PHARMACIST_DB = {
    "ภญ.มนรดา พิพรรธนัชกุล(INV)": "warehouse",
    "ภญ.นันทาศิริ แก้วพันสี(INV)": "warehouse",
    "ภญ.จตุพร สุมิตสวรรค์(INV)": "warehouse",
    "ภญ.วัชรินทร์ ถาวโรภาส": "general",
    "ภญ.กัณฑิมา ดาระอินทร์": "general",
    "ภญ.มนัสนันท์ เกียรติสมบูลย์": "general",
    "ภก.พุทธิวัฒน์ เชื้อชาติทางธกุล": "general",
    "ภก.พัทยา สุ่มแก้ว": "general",
    "ภญ.พลอยไพลิน ศรีม่วง": "general",
    "ภก.ชัยพงศ์ ดำรงสุวรรณ": "general",
    "ภญ.กาญจนา คนเที่ยง": "general",
    "ภญ.โสมิกา คล่ำเงิน": "general",
    "ภญ.สภาวัลย์ อภิชาติตรากูล": "general",
    "ภก.กัมพล เกตุสุวรรณ": "general",
    "ภก.พันธวงศ์ โรจน์ธนศิริวนิช": "general",
    "ภญ.เบญญาศิริ รุ่งสว่าง": "general",
    "ภก.ธีรพร เมฆพัฒน์": "general",
    "ภญ.กมลพัชร เอื้อกุศลสมบูรณ์": "general",
    "ภก.เสถียรพงศ์ แก้วเมธีกุล": "general",
    "ภก.ศุภณัฐ จินดาขัด": "general",
}

# ==========================================
# 🔐 เลือกโหมดเข้าสู่ระบบ: ผู้ปฏิบัติงาน / System Admin
# ==========================================
login_mode = st.sidebar.radio(
    "โหมดเข้าสู่ระบบ",
    ["👩‍⚕️ ผู้ปฏิบัติงาน", "🛠️ System Admin"],
    key="login_mode",
)

# System Admin เป็นบัญชีของผู้พัฒนา/ผู้ดูแลโปรแกรม แยกจากเภสัชกรคลังยา
# ใช้เฉพาะจัดการ credential เช่น reset รหัสผ่าน ไม่ใช้ทำรายการยา
if login_mode == "🛠️ System Admin":
    st.sidebar.markdown("---")
    st.sidebar.markdown("**🛠️ ผู้ดูแลระบบ (System Admin)**")

    configured_system_admin_password = get_system_admin_password()
    if not configured_system_admin_password:
        st.error(
            "⚠️ ยังไม่ได้ตั้งรหัส System Admin ใน Streamlit Secrets "
            "กรุณาเพิ่มค่า SYSTEM_ADMIN_PASSWORD ก่อนใช้งานเมนูนี้"
        )
        st.code('SYSTEM_ADMIN_PASSWORD = "ใส่รหัสที่คุณตั้งเองตรงนี้"', language="toml")
        st.caption(
            "รหัสนี้จะไม่ถูกเขียนไว้ใน app.py และไม่ถูกเก็บใน Google Sheets"
        )
        st.stop()

    system_admin_authenticated = st.session_state.get("system_admin_authenticated", False)

    if not system_admin_authenticated:
        system_admin_password = st.sidebar.text_input(
            "รหัส System Admin",
            type="password",
            autocomplete="off",
            key="system_admin_password_input",
        )
        if st.sidebar.button("🔓 เข้าสู่ System Admin", key="system_admin_login_btn"):
            if verify_system_admin_password(system_admin_password):
                st.session_state["system_admin_authenticated"] = True
                st.rerun()
            else:
                st.sidebar.error("❌ รหัส System Admin ไม่ถูกต้อง")
        st.stop()

    st.sidebar.success("✅ เข้าสู่ระบบ System Admin")
    if st.sidebar.button("🚪 ออกจาก System Admin", key="system_admin_logout_btn"):
        st.session_state.pop("system_admin_authenticated", None)
        st.rerun()

    st.subheader("🛠️ จัดการบัญชีเจ้าหน้าที่คลังยา")
    st.info(
        "System Admin ใช้สำหรับดูสถานะบัญชีและรีเซ็ตรหัสผ่านของเจ้าหน้าที่คลังยาเท่านั้น "
        "ไม่ใช้ทำรายการยาหรือจัดการสถานะยืม-คืนแทนเจ้าหน้าที่คลัง"
    )

    warehouse_users = [
        name for name, role in PHARMACIST_DB.items() if role == "warehouse"
    ]
    selected_warehouse_user = st.selectbox(
        "เลือกเจ้าหน้าที่คลังยาที่ต้องการจัดการ",
        warehouse_users,
        key="system_admin_selected_warehouse",
    )

    selected_credential = get_latest_admin_credential(selected_warehouse_user)
    if selected_credential:
        st.success(f"✅ {selected_warehouse_user} มีรหัสผ่านสำหรับเข้าใช้งานแล้ว")
    else:
        st.warning(
            f"⚠️ {selected_warehouse_user} ยังไม่ได้ตั้งรหัสผ่าน หรือบัญชีถูกรีเซ็ตแล้ว"
        )

    st.markdown("#### รีเซ็ตรหัสผ่าน")
    st.write(
        "เมื่อรีเซ็ตแล้ว ระบบจะไม่เปิดเผยหรือกู้รหัสเดิม "
        "เจ้าหน้าที่คนนี้จะต้องตั้งรหัสผ่านใหม่ด้วยตนเองเมื่อเข้าใช้งานครั้งถัดไป"
    )
    confirm_reset = st.checkbox(
        f"ยืนยันว่าต้องการรีเซ็ตรหัสของ {selected_warehouse_user}",
        key="system_admin_confirm_reset",
    )

    if st.button(
        "🔄 รีเซ็ตรหัสผ่านเจ้าหน้าที่คลัง",
        disabled=not confirm_reset,
        type="primary",
        key="system_admin_reset_btn",
    ):
        ok, msg = reset_warehouse_password(selected_warehouse_user)
        if ok:
            # ถ้า System Admin กับเจ้าหน้าที่คลังเปิดอยู่ใน session เดียวกัน ให้บังคับ login ใหม่
            st.session_state.pop(
                f"warehouse_authenticated::{selected_warehouse_user}", None
            )
            st.success(
                f"รีเซ็ตรหัสของ {selected_warehouse_user} สำเร็จ "
                "ครั้งถัดไปเจ้าหน้าที่จะต้องตั้งรหัสใหม่"
            )
            st.rerun()
        else:
            st.error(f"ไม่สามารถรีเซ็ตรหัสผ่านได้: {msg}")

    st.stop()

# --- โหมดผู้ปฏิบัติงาน ---
pharmacist_names = ["เลือกชื่อ..."] + list(PHARMACIST_DB.keys())
user_name = st.sidebar.selectbox("ผู้ทำรายการ", pharmacist_names)

st.sidebar.markdown("---")
st.sidebar.header("📌 เลือกเมนูทำงาน")

# บังคับให้เลือกชื่อก่อน ถึงจะเห็นเมนู
if user_name == "เลือกชื่อ...":
    st.warning("⚠️ กรุณาเลือก 'ผู้ทำรายการ' ที่เมนูด้านซ้ายมือก่อนเริ่มทำงานค่ะ")
    st.stop()

# 2. ตรวจสอบสิทธิ์การเข้าถึง (ดึงสิทธิ์จากฐานข้อมูล)
user_role = PHARMACIST_DB.get(user_name, "general")

st.sidebar.markdown("---")
# ==========================================
# 🔐 ระบบ Security — รหัสผ่านเฉพาะเจ้าหน้าที่คลังยา
# ==========================================
warehouse_authenticated = False

if user_role == "warehouse":
    st.sidebar.markdown("**🔐 ยืนยันตัวตน (เจ้าหน้าที่คลังยา)**")

    auth_state_key = f"warehouse_authenticated::{user_name}"
    credential = get_latest_admin_credential(user_name)

    # --- ครั้งแรก: ให้เจ้าหน้าที่คลังยาตั้งรหัสผ่านของตัวเอง ---
    if credential is None:
        st.sidebar.info("บัญชีนี้ยังไม่ได้ตั้งรหัสผ่าน กรุณาตั้งรหัสผ่านสำหรับใช้งานครั้งแรก")
        new_password = st.sidebar.text_input(
            "ตั้งรหัสผ่านใหม่",
            type="password",
            autocomplete="off",
            key=f"setup_password::{user_name}",
            help="อย่างน้อย 6 ตัวอักษร",
        )
        confirm_password = st.sidebar.text_input(
            "ยืนยันรหัสผ่านใหม่",
            type="password",
            autocomplete="off",
            key=f"setup_password_confirm::{user_name}",
        )

        if st.sidebar.button("✅ ตั้งรหัสผ่าน", key=f"setup_password_btn::{user_name}"):
            if not password_is_valid(new_password):
                st.sidebar.error("รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร")
            elif new_password != confirm_password:
                st.sidebar.error("รหัสผ่านทั้งสองช่องไม่ตรงกัน")
            else:
                ok, msg = set_admin_password(user_name, new_password)
                if ok:
                    st.session_state[auth_state_key] = True
                    st.sidebar.success("ตั้งรหัสผ่านสำเร็จ")
                    st.rerun()
                else:
                    st.sidebar.error(
                        "ไม่สามารถบันทึกรหัสผ่านได้ กรุณาตรวจสอบว่ามีชีต "
                        f"'{ADMIN_PASSWORD_SHEET}' และ Google Apps Script อนุญาตให้บันทึกชีตนี้"
                    )
        st.stop()

    # --- มีรหัสแล้ว: ตรวจสอบ session ก่อนถามรหัสซ้ำ ---
    warehouse_authenticated = st.session_state.get(auth_state_key, False)

    if not warehouse_authenticated:
        admin_password = st.sidebar.text_input(
            "รหัสผ่านส่วนตัว",
            type="password",
            autocomplete="off",
            key=f"login_password::{user_name}",
        )
        if st.sidebar.button("🔓 เข้าสู่ระบบ", key=f"login_btn::{user_name}"):
            if verify_admin_password(user_name, admin_password):
                st.session_state[auth_state_key] = True
                st.rerun()
            else:
                st.sidebar.error("❌ รหัสผ่านไม่ถูกต้อง")
        st.stop()

    warehouse_authenticated = True
    st.sidebar.success("✅ เข้าสู่ระบบระดับคลังยา")

    # เปลี่ยนรหัสผ่านได้ด้วยตนเองหลังเข้าสู่ระบบแล้ว
    with st.sidebar.expander("🔑 เปลี่ยนรหัสผ่าน"):
        current_password = st.text_input(
            "รหัสผ่านปัจจุบัน",
            type="password",
            autocomplete="off",
            key=f"change_current::{user_name}",
        )
        changed_password = st.text_input(
            "รหัสผ่านใหม่",
            type="password",
            autocomplete="off",
            key=f"change_new::{user_name}",
            help="อย่างน้อย 6 ตัวอักษร",
        )
        changed_password_confirm = st.text_input(
            "ยืนยันรหัสผ่านใหม่",
            type="password",
            autocomplete="off",
            key=f"change_confirm::{user_name}",
        )
        if st.button("บันทึกรหัสผ่านใหม่", key=f"change_btn::{user_name}"):
            if not verify_admin_password(user_name, current_password):
                st.error("รหัสผ่านปัจจุบันไม่ถูกต้อง")
            elif not password_is_valid(changed_password):
                st.error("รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัวอักษร")
            elif changed_password != changed_password_confirm:
                st.error("รหัสผ่านใหม่ทั้งสองช่องไม่ตรงกัน")
            elif hmac.compare_digest(current_password, changed_password):
                st.error("รหัสผ่านใหม่ต้องไม่เหมือนรหัสผ่านเดิม")
            else:
                ok, msg = set_admin_password(user_name, changed_password)
                if ok:
                    st.success("เปลี่ยนรหัสผ่านสำเร็จ")
                else:
                    st.error(f"ไม่สามารถเปลี่ยนรหัสผ่านได้: {msg}")

    if st.sidebar.button("🚪 ออกจากระบบคลังยา", key=f"logout_btn::{user_name}"):
        st.session_state.pop(auth_state_key, None)
        st.rerun()
else:
    st.sidebar.success("✅ เข้าสู่ระบบระดับ User")

st.sidebar.markdown("---")


# เมนูพื้นฐานที่ทุกคนเห็น
menu_options = [
    "1. จ่ายยาออก (Refer รพช.)", 
    "2. ยืมยาเข้า (ยา รพ. เราไม่พอ)", 
    "3. ให้ รพ.อื่นยืมยา (ยา รพ.อื่นขาด)", 
    "4. 📊 Dashboard สรุปข้อมูล"
]

# 3. ถ้าเป็นเภสัชคลังยา (warehouse) และผ่านรหัสผ่าน ให้เพิ่มเมนูที่ 5 เข้าไป
if user_role == "warehouse" and warehouse_authenticated:
    menu_options.append("5. 🔄 ติดตามสถานะคลังยา (Admin)")

# แสดงเมนูตามสิทธิ์
menu = st.sidebar.radio("เลือกกรณีที่ต้องการ", menu_options)


# ==========================================
# เมนูที่ 1: จ่ายยาออก (Refer Out)
# ==========================================
if menu == "1. จ่ายยาออก (Refer รพช.)":
    st.subheader("📤 กรณีที่ 1: จ่ายยาให้ผู้ป่วย Refer กลับ รพช.")

    refer_items = get_item_list("refer_items")
    doc_no = get_transaction_no("refer_doc_no", "REF")

    col1, col2 = st.columns(2)
    with col1:
        pt_name = st.text_input("ชื่อ-นามสกุล ผู้ป่วย")
        hn = st.text_input("รหัสประจำตัวผู้ป่วย (HN)")
        target_hosp = st.selectbox("รพช. ปลายทาง", ["ทุ่งเสลี่ยม", "ศรีสัชนาลัย", "ศรีนคร", "สวรรคโลก", "อื่นๆ"])
    with col2:
        date_out = st.date_input("วันที่ทำรายการ", datetime.date.today())
        st.text_input("เลขที่ใบยืม (สร้างอัตโนมัติ)", value=doc_no, disabled=True, key=f"refer_doc_display_{doc_no}")
        patient_type = st.radio("ประเภทผู้ป่วย", ["OPD", "IPD / อื่นๆ"], horizontal=True, key="refer_patient_type")

    previous_refer_mode = st.session_state.get("refer_items_mode")
    if previous_refer_mode is None:
        st.session_state["refer_items_mode"] = patient_type
    elif previous_refer_mode != patient_type:
        st.session_state["refer_items"] = []
        refer_items = st.session_state["refer_items"]
        st.session_state["refer_items_mode"] = patient_type
        st.info("ℹ️ เปลี่ยนประเภทผู้ป่วยแล้ว รายการยาของประเภทเดิมถูกล้างเพื่อป้องกันข้อมูลปะปน")

    is_opd = patient_type == "OPD"
    if is_opd:
        st.info(
            "🏥 **Flow OPD Refer Back:** รพ.เราจ่ายยาให้ผู้ป่วยสำหรับ 3 วัน "
            "และจำนวนที่เหลือจะบันทึกเป็นรายการที่ รพ.ปลายทางต้องทำเรื่องยืมยา"
        )

    st.markdown("**เพิ่มรายการยา/เวชภัณฑ์**")

    # ใช้ widget ปกติแทน st.form เพื่อให้ช่อง "หน่วยอื่นๆ" แสดงเฉพาะเมื่อเลือก "อื่นๆ"
    # ทำให้การกรอกตามปกติอยู่ในบรรทัดเดียวกัน และยังเปลี่ยน UI ได้ทันทีเมื่อเลือกหน่วย
    refer_input_keys = [
        "refer_drug_name_input", "refer_qty_3day_input", "refer_borrow_qty_input",
        "refer_qty_input", "refer_unit_choice_input", "refer_custom_unit_input",
    ]
    if st.session_state.pop("reset_refer_item_inputs", False):
        for key in refer_input_keys:
            st.session_state.pop(key, None)

    if is_opd:
        col_d1, col_d2, col_d3, col_d4 = st.columns([2.4, 1.25, 1.55, 1.0])
        with col_d1:
            drug_name = st.text_input("ชื่อยา หรือ เวชภัณฑ์", key="refer_drug_name_input")
        with col_d2:
            qty_3day = st.number_input(
                "จ่าย 3 วัน", min_value=0.0, step=1.0, key="refer_qty_3day_input"
            )
        with col_d3:
            borrow_qty = st.number_input(
                "รพ.ปลายทางยืม", min_value=0.0, step=1.0, key="refer_borrow_qty_input"
            )
        with col_d4:
            unit_choice = st.selectbox("หน่วย", UNITS, key="refer_unit_choice_input")

        # แสดงช่องระบุหน่วยเพิ่มเติมเฉพาะกรณีเลือก "อื่นๆ" เท่านั้น
        custom_unit = ""
        if unit_choice == "อื่นๆ":
            custom_unit = st.text_input(
                "ระบุหน่วยอื่นๆ", key="refer_custom_unit_input", placeholder="เช่น ซอง"
            )

        add_refer_item = st.button("➕ เพิ่มรายการยา", key="add_refer_item_opd")
        if add_refer_item:
            unit = custom_unit.strip() if unit_choice == "อื่นๆ" else unit_choice
            if not drug_name.strip():
                st.error("กรุณาระบุชื่อยา/เวชภัณฑ์")
            elif not unit:
                st.error("กรุณาระบุหน่วย")
            elif qty_3day <= 0 and borrow_qty <= 0:
                st.error("กรุณาระบุจำนวนยาอย่างน้อย 1 ช่อง")
            else:
                refer_items.append({
                    "drug_name": drug_name.strip(),
                    "qty_3day": qty_3day,
                    "borrow_qty": borrow_qty,
                    "unit": unit,
                })
                st.session_state["reset_refer_item_inputs"] = True
                st.rerun()
    else:
        col_d1, col_d2, col_d3 = st.columns([2.5, 1.0, 1.0])
        with col_d1:
            drug_name = st.text_input("ชื่อยา หรือ เวชภัณฑ์", key="refer_drug_name_input")
        with col_d2:
            qty = st.number_input("จำนวน", min_value=1.0, step=1.0, key="refer_qty_input")
        with col_d3:
            unit_choice = st.selectbox("หน่วย", UNITS, key="refer_unit_choice_input")

        custom_unit = ""
        if unit_choice == "อื่นๆ":
            custom_unit = st.text_input(
                "ระบุหน่วยอื่นๆ", key="refer_custom_unit_input", placeholder="เช่น ซอง"
            )

        add_refer_item = st.button("➕ เพิ่มรายการยา", key="add_refer_item_non_opd")
        if add_refer_item:
            unit = custom_unit.strip() if unit_choice == "อื่นๆ" else unit_choice
            if not drug_name.strip():
                st.error("กรุณาระบุชื่อยา/เวชภัณฑ์")
            elif not unit:
                st.error("กรุณาระบุหน่วย")
            else:
                refer_items.append({
                    "drug_name": drug_name.strip(),
                    "qty": qty,
                    "unit": unit,
                })
                st.session_state["reset_refer_item_inputs"] = True
                st.rerun()

    if refer_items:
        st.markdown(f"**รายการในใบนี้: {len(refer_items)} รายการ**")
        for idx, item in enumerate(refer_items):
            c1, c2 = st.columns([8, 1])
            with c1:
                if is_opd:
                    st.write(
                        f"{idx + 1}. {item['drug_name']} | จ่าย 3 วัน: {format_qty(item['qty_3day'])} {item['unit']} "
                        f"| รพ.ปลายทางยืม: {format_qty(item['borrow_qty'])} {item['unit']}"
                    )
                else:
                    st.write(f"{idx + 1}. {item['drug_name']} | {format_qty(item['qty'])} {item['unit']}")
            with c2:
                if st.button("🗑️", key=f"remove_refer_{idx}", help="ลบรายการนี้"):
                    refer_items.pop(idx)
                    st.rerun()
    else:
        st.caption("ยังไม่มีรายการยาในใบนี้")

    st.markdown("**เหตุผลความจำเป็น**")
    reason_choice = st.radio(
        "เลือกเหตุผล:",
        ["Refer Back", "ผู้ป่วยฉุกเฉิน / อุบัติเหตุ", "เหตุผลอื่นๆ ...."],
        horizontal=True,
        label_visibility="collapsed",
    )
    note = st.text_input("โปรดระบุเหตุผลอื่นๆ ...") if reason_choice == "เหตุผลอื่นๆ ...." else reason_choice

    st.markdown("---")
    if st.button("💾 บันทึกข้อมูล และ สร้างใบให้ยืมยา", type="primary"):
        if not pt_name.strip() or not hn.strip():
            st.error("กรุณากรอกชื่อผู้ป่วยและ HN ให้ครบ")
        elif not refer_items:
            st.error("กรุณาเพิ่มรายการยาอย่างน้อย 1 รายการ")
        else:
            rows_to_save = []
            for idx, item in enumerate(refer_items, start=1):
                line_doc_no = make_line_doc_no(doc_no, idx)
                if is_opd:
                    # เฉพาะส่วนที่เหลือหลังจ่าย 3 วันเท่านั้นที่เป็นรายการยืมจาก รพ.ปลายทาง
                    row_note = (
                        f"{note} | OPD Refer Back | จ่ายผู้ป่วย 3 วัน {format_qty(item['qty_3day'])} {item['unit']} "
                        f"| รพ.ปลายทางต้องยืม {format_qty(item['borrow_qty'])} {item['unit']} | ใบหลัก {doc_no}"
                    )
                    if item["borrow_qty"] > 0:
                        rows_to_save.append([
                            line_doc_no, str(date_out), user_name, target_hosp, hn,
                            item["drug_name"], item["borrow_qty"], item["unit"], "-", row_note,
                            "รอเอกสารยืม"
                        ])
                else:
                    row_note = f"{note} | ใบหลัก {doc_no}"
                    rows_to_save.append([
                        line_doc_no, str(date_out), user_name, target_hosp, hn,
                        item["drug_name"], item["qty"], item["unit"], "-", row_note,
                        "รอคืนยา"
                    ])

            saved_count = 0
            errors = []
            with st.spinner('กำลังบันทึกข้อมูลลง Google Sheets...'):
                for row_data in rows_to_save:
                    is_saved, debug_msg = save_to_google_sheets("Outbound_Refer", row_data=row_data, action="append")
                    if is_saved:
                        saved_count += 1
                    else:
                        errors.append(f"{row_data[0]}: {debug_msg}")

            if errors:
                st.error("❌ บันทึกข้อมูลไม่ครบทุกรายการ\n" + "\n".join(errors))
            else:
                if rows_to_save:
                    st.success(f"✅ บันทึกข้อมูลสำเร็จ {saved_count} รายการ ภายใต้ใบหลัก {doc_no}")
                else:
                    st.success("✅ OPD รายนี้ไม่มีส่วนยาคงเหลือที่ต้องทำเรื่องยืม จึงไม่มีรายการยืมที่ต้องบันทึก")

                try:
                    doc_refer = DocxTemplate("template_refer_out.docx")
                    context_refer = {
                        'target_hospital': target_hosp,
                        'pt_name': pt_name,
                        'hn': hn,
                        # ใช้ marker ชั่วคราว แล้วแทนด้วยตารางไร้เส้นหลัง render
                        'drug_details': '__DRUG_TABLE__',
                        'user_name': user_name,
                        'thai_date': get_thai_date(date_out),
                        # ตัวแปรเสริม ใช้ได้ทันทีหากเพิ่ม placeholder ใน template ภายหลัง
                        'patient_type': patient_type,
                        'doc_no': doc_no,
                    }
                    doc_refer.render(context_refer)
                    bio_refer = io.BytesIO()
                    doc_refer.save(bio_refer)
                    bio_refer = replace_drug_marker_with_table(bio_refer, refer_items, opd_mode=is_opd)

                    st.download_button(
                        label="📥 โหลดใบให้ยืมยา (ส่งตัวผู้ป่วย)",
                        data=bio_refer.getvalue(),
                        file_name=f"ReferOut_{doc_no}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"⚠️ เกิดข้อผิดพลาดในการสร้างเอกสาร: {e}")

                # เตรียมเลขเอกสารและรายการใหม่สำหรับ transaction ถัดไปเมื่อหน้าจอ rerun
                st.session_state["refer_items"] = []
                st.session_state["refer_doc_no"] = f"REF-{datetime.datetime.now().strftime('%y%m%d-%H%M%S')}"


# ==========================================
# เมนูที่ 2: ยืมยาเข้า (Borrow In)
# ==========================================
elif menu == "2. ยืมยาเข้า (ยา รพ. เราไม่พอ)":
    st.subheader("📥 กรณีที่ 2: ยืมยาจาก รพ. อื่น (ยาขาดคลัง)")

    borrow_items = get_item_list("borrow_items")
    borrow_no = get_transaction_no("borrow_doc_no", "REQ")

    col1, col2 = st.columns(2)
    with col1:
        st.text_input("เลขที่การยืม (สร้างอัตโนมัติ)", value=borrow_no, disabled=True, key=f"borrow_doc_display_{borrow_no}")
        source_hosp = st.text_input("รพ. ที่เราต้องการขอยืม")
    with col2:
        date_req = st.date_input("วันที่แจ้งเรื่อง", datetime.date.today())

    st.markdown("**เพิ่มรายการยาที่ต้องการยืม**")
    with st.form("borrow_add_item_form", clear_on_submit=True):
        col_d1, col_d2, col_d3 = st.columns([2, 1, 1])
        with col_d1:
            drug_missing = st.text_input("ชื่อยาที่ขาด/ต้องการยืม")
        with col_d2:
            borrow_qty = st.number_input("จำนวนที่ต้องการยืม", min_value=1.0, step=1.0)
        with col_d3:
            unit_choice2 = st.selectbox("หน่วย (ยืมเข้า)", UNITS)
            custom_unit2 = st.text_input("หน่วยอื่นๆ (กรอกเมื่อเลือก 'อื่นๆ')")
        add_borrow_item = st.form_submit_button("➕ เพิ่มรายการยา")

        if add_borrow_item:
            unit2 = custom_unit2.strip() if unit_choice2 == "อื่นๆ" else unit_choice2
            if not drug_missing.strip():
                st.error("กรุณาระบุชื่อยา")
            elif not unit2:
                st.error("กรุณาระบุหน่วย")
            else:
                borrow_items.append({"drug_name": drug_missing.strip(), "qty": borrow_qty, "unit": unit2})
                st.rerun()

    if borrow_items:
        st.markdown(f"**รายการในใบนี้: {len(borrow_items)} รายการ**")
        for idx, item in enumerate(borrow_items):
            c1, c2 = st.columns([8, 1])
            with c1:
                st.write(f"{idx + 1}. {item['drug_name']} | {format_qty(item['qty'])} {item['unit']}")
            with c2:
                if st.button("🗑️", key=f"remove_borrow_{idx}", help="ลบรายการนี้"):
                    borrow_items.pop(idx)
                    st.rerun()
    else:
        st.caption("ยังไม่มีรายการยาในใบนี้")

    st.markdown("---")
    if st.button("💾 บันทึกข้อมูลการยืมยาเข้าคลัง", type="primary"):
        if not source_hosp.strip():
            st.error("กรุณาระบุโรงพยาบาลที่ต้องการขอยืม")
        elif not borrow_items:
            st.error("กรุณาเพิ่มรายการยาอย่างน้อย 1 รายการ")
        else:
            saved_count = 0
            errors = []
            with st.spinner('กำลังบันทึกข้อมูลลง Google Sheets...'):
                for idx, item in enumerate(borrow_items, start=1):
                    line_doc_no = make_line_doc_no(borrow_no, idx)
                    row_data_in = [
                        line_doc_no, str(date_req), user_name, item["drug_name"],
                        item["qty"], item["unit"], source_hosp, "รอคืนยา", f"ใบหลัก {borrow_no}"
                    ]
                    is_saved, debug_msg = save_to_google_sheets("Inbound_Shortage", row_data=row_data_in, action="append")
                    if is_saved:
                        saved_count += 1
                    else:
                        errors.append(f"{line_doc_no}: {debug_msg}")

            if errors:
                st.error("❌ บันทึกข้อมูลไม่ครบทุกรายการ\n" + "\n".join(errors))
            else:
                st.success(f"✅ บันทึกข้อมูลสำเร็จ {saved_count} รายการ ภายใต้ใบหลัก {borrow_no}")

    st.markdown("---")
    st.subheader("🖨️ พิมพ์เอกสารขอยืมยาเข้า รพ.")
    col_memo, col_official = st.columns(2)

    with col_memo:
        st.info("🌙 สำหรับเภสัชกรอยู่เวร")
        if st.button("📄 พิมพ์บันทึกข้อความ"):
            if not borrow_items:
                st.error("กรุณาเพิ่มรายการยาอย่างน้อย 1 รายการก่อนพิมพ์เอกสาร")
            else:
                try:
                    doc_memo = DocxTemplate("template_memo.docx")
                    context_memo = {
                        'target_hospital': source_hosp,
                        'drug_details': drug_details_text(borrow_items),
                        'user_name': user_name,
                        'thai_date': get_thai_date(date_req),
                        'doc_no': borrow_no,
                    }
                    doc_memo.render(context_memo)
                    bio_memo = io.BytesIO()
                    doc_memo.save(bio_memo)
                    st.download_button(
                        "📥 โหลดบันทึกข้อความ",
                        data=bio_memo.getvalue(),
                        file_name=f"Memo_{borrow_no}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"⚠️ เกิดข้อผิดพลาด: {e}")

    with col_official:
        st.success("☀️ สำหรับคลังยา (งานสารบรรณ)")
        formal_reason = st.selectbox(
            "เลือกเหตุผลในหนังสือ:",
            ["ยาขาดชั่วคราว จำเป็นต้องใช้เร่งด่วน", "ไม่มียาในบัญชีของโรงพยาบาล", "เป็นยา จ2 ไม่มียาในบัญชียา"]
        )
        if st.button("🦅 พิมพ์หนังสือตราครุฑ"):
            if not borrow_items:
                st.error("กรุณาเพิ่มรายการยาอย่างน้อย 1 รายการก่อนพิมพ์เอกสาร")
            else:
                try:
                    doc_official = DocxTemplate("template_official.docx")
                    context_official = {
                        'target_hospital': source_hosp,
                        'reason': formal_reason,
                        'drug_details': drug_details_text(borrow_items),
                        'thai_date': get_thai_date(date_req),
                        'doc_no': borrow_no,
                    }
                    doc_official.render(context_official)
                    bio_official = io.BytesIO()
                    doc_official.save(bio_official)
                    st.download_button(
                        "📥 โหลดหนังสือตราครุฑ",
                        data=bio_official.getvalue(),
                        file_name=f"Official_{borrow_no}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"⚠️ เกิดข้อผิดพลาด: {e}")


# ==========================================
# เมนูที่ 3: ให้ รพ.อื่นยืมยา (Lend Out)
# ==========================================
elif menu == "3. ให้ รพ.อื่นยืมยา (ยา รพ.อื่นขาด)":
    st.subheader("📤 กรณีที่ 3: ให้ รพ.อื่นยืมยา (ยา รพ.อื่นขาดคลัง)")

    lend_items = get_item_list("lend_items")
    lend_no = get_transaction_no("lend_doc_no", "LEND")

    col1, col2 = st.columns(2)
    with col1:
        st.text_input("เลขที่การให้ยืม (สร้างอัตโนมัติ)", value=lend_no, disabled=True, key=f"lend_doc_display_{lend_no}")
    with col2:
        date_lend = st.date_input("วันที่ให้ยืม", datetime.date.today())

    target_hosp_lend = st.text_input("ชื่อ รพ. ที่มายืมยา")

    st.markdown("**เพิ่มรายการยาที่ให้ยืม**")
    with st.form("lend_add_item_form", clear_on_submit=True):
        col_d1, col_d2, col_d3 = st.columns([2, 1, 1])
        with col_d1:
            drug_lended = st.text_input("ชื่อยา หรือ เวชภัณฑ์ที่ให้ยืม")
        with col_d2:
            lend_qty = st.number_input("จำนวนที่ให้ยืม", min_value=1.0, step=1.0)
        with col_d3:
            unit_choice_lend = st.selectbox("หน่วย", UNITS)
            custom_unit_lend = st.text_input("หน่วยอื่นๆ (กรอกเมื่อเลือก 'อื่นๆ')")
        add_lend_item = st.form_submit_button("➕ เพิ่มรายการยา")

        if add_lend_item:
            unit_lend = custom_unit_lend.strip() if unit_choice_lend == "อื่นๆ" else unit_choice_lend
            if not drug_lended.strip():
                st.error("กรุณาระบุชื่อยา/เวชภัณฑ์")
            elif not unit_lend:
                st.error("กรุณาระบุหน่วย")
            else:
                lend_items.append({"drug_name": drug_lended.strip(), "qty": lend_qty, "unit": unit_lend})
                st.rerun()

    if lend_items:
        st.markdown(f"**รายการในใบนี้: {len(lend_items)} รายการ**")
        for idx, item in enumerate(lend_items):
            c1, c2 = st.columns([8, 1])
            with c1:
                st.write(f"{idx + 1}. {item['drug_name']} | {format_qty(item['qty'])} {item['unit']}")
            with c2:
                if st.button("🗑️", key=f"remove_lend_{idx}", help="ลบรายการนี้"):
                    lend_items.pop(idx)
                    st.rerun()
    else:
        st.caption("ยังไม่มีรายการยาในใบนี้")

    st.markdown("---")
    if st.button("💾 บันทึกข้อมูลให้ รพ.อื่นยืมยา", type="primary"):
        if not target_hosp_lend.strip():
            st.error("กรุณาระบุชื่อโรงพยาบาลที่มายืมยา")
        elif not lend_items:
            st.error("กรุณาเพิ่มรายการยาอย่างน้อย 1 รายการ")
        else:
            saved_count = 0
            errors = []
            with st.spinner('กำลังบันทึกข้อมูลลง Google Sheets...'):
                for idx, item in enumerate(lend_items, start=1):
                    line_doc_no = make_line_doc_no(lend_no, idx)
                    row_data_lend = [
                        line_doc_no, str(date_lend), user_name, target_hosp_lend,
                        item["drug_name"], item["qty"], item["unit"], "รอรับคืน"
                    ]
                    is_saved, debug_msg = save_to_google_sheets("Outbound_Lend", row_data=row_data_lend, action="append")
                    if is_saved:
                        saved_count += 1
                    else:
                        errors.append(f"{line_doc_no}: {debug_msg}")

            if errors:
                st.error("❌ บันทึกข้อมูลไม่ครบทุกรายการ\n" + "\n".join(errors))
            else:
                st.success(f"✅ บันทึกข้อมูลสำเร็จ {saved_count} รายการ ภายใต้ใบหลัก {lend_no}")
                st.session_state["lend_items"] = []
                st.session_state["lend_doc_no"] = f"LEND-{datetime.datetime.now().strftime('%y%m%d-%H%M%S')}"

# ==========================================
# เมนูที่ 4: Dashboard สรุปข้อมูล
# ==========================================
elif menu == "4. 📊 Dashboard สรุปข้อมูล":
    st.subheader("📊 Dashboard สรุปสถานะการยืม-คืนยา")
    
    # ✨ เพิ่มข้อความแจ้งเตือนผู้ใช้งานตรงนี้
    st.info("💡 **ข้อแนะนำ:** หากข้อมูลในตารางยังไม่อัปเดต ให้กดปุ่ม **'🔄 อัปเดตข้อมูลล่าสุด'** หรือกดเลือกช่วงวันที่ ")
    
    # ⚠️ นำลิงก์ Embed ของแต่ละหน้าจาก Looker Studio มาใส่ตรงนี้ (อย่าลืมเติม embed/ นะคะ)
    URL_REFER_OUT = "https://datastudio.google.com/embed/reporting/eb15729f-a596-4adc-bb19-b28269e6838d/page/g566F"
    URL_BORROW_IN = "https://datastudio.google.com/embed/reporting/eb15729f-a596-4adc-bb19-b28269e6838d/page/p_7tlqc5pq6d"
    URL_LEND_OUT  = "https://datastudio.google.com/embed/reporting/eb15729f-a596-4adc-bb19-b28269e6838d/page/p_rq3f76pq6d"
    
    # สร้าง Tabs เพื่อแยกหน้า Dashboard ให้ดูง่าย
    dash_tab1, dash_tab2, dash_tab3 = st.tabs(["📤 จ่ายยาออก (Refer Out)", "📥 เรายืม รพ.อื่น (Borrow In)", "🤝 รพ.อื่นยืมเรา (Lend Out)"])
    
    with dash_tab1:
        st.components.v1.iframe(URL_REFER_OUT, width=800, height=700, scrolling=True)
        
    with dash_tab2:
        st.components.v1.iframe(URL_BORROW_IN, width=800, height=700, scrolling=True)
        
    with dash_tab3:
        st.components.v1.iframe(URL_LEND_OUT, width=800, height=700, scrolling=True)

# ==========================================
# เมนูที่ 5: ติดตามสถานะ (รับคืน/ส่งคืน) - เฉพาะ Admin
# ==========================================
elif menu == "5. 🔄 ติดตามสถานะคลังยา (Admin)":
    st.subheader("🔄 จัดการสถานะการยืม-คืนยา")
    
    tab1, tab2, tab3 = st.tabs(["📥 OPD / รับยาคืน (Refer Out)", "📤 ส่งยาคืน (เรายืมเขา)", "📥 รับยาคืน (รพ.อื่นยืมเรา)"])
    
    # ---------------- แท็บที่ 1: OPD รอเอกสารยืม / รับยาคืน ----------------
    with tab1:
        out_data = get_from_google_sheets("Outbound_Refer")
        
        if out_data and len(out_data) > 1:
            waiting_docs = [row for row in out_data[1:] if len(row) > 10 and row[10] == "รอเอกสารยืม"]
            pending_out = [row for row in out_data[1:] if len(row) > 10 and row[10] == "รอคืนยา"]

            st.markdown("**1) OPD Refer Back: รายการที่รอ รพ.ปลายทางส่งเรื่องยืม**")
            if waiting_docs:
                doc_options = [f"เลขที่: {row[0]} | รพ: {row[3]} | ยา: {row[5]} ({row[6]} {row[7]})" for row in waiting_docs]
                selected_doc = st.selectbox(
                    "เลือกรายการที่ได้รับเอกสารยืมจาก รพ.ปลายทางแล้ว:",
                    ["-- เลือกรายการ --"] + doc_options,
                    key="opd_waiting_doc_select"
                )
                if selected_doc != "-- เลือกรายการ --":
                    doc_id_waiting = selected_doc.split(" | ")[0].replace("เลขที่: ", "")
                    if st.button("📄 ยืนยันได้รับเรื่องยืมแล้ว", key="confirm_borrow_doc"):
                        with st.spinner("กำลังอัปเดตฐานข้อมูล..."):
                            is_saved, msg = save_to_google_sheets(
                                "Outbound_Refer", action="update", doc_id=doc_id_waiting, new_status="รอคืนยา"
                            )
                            if is_saved:
                                st.success(f"✅ {doc_id_waiting} เปลี่ยนสถานะเป็น 'รอคืนยา' แล้ว")
                                st.rerun()
                            else:
                                st.error(f"❌ ผิดพลาด: {msg}")
            else:
                st.info("✨ ไม่มีรายการ OPD ที่รอเอกสารยืม")

            st.markdown("---")
            st.markdown("**2) รายการที่จ่าย/ให้ยืมออกไป และยังไม่ได้รับคืน**")
            if pending_out:
                out_options = [f"เลขที่: {row[0]} | รพ: {row[3]} | ยา: {row[5]} ({row[6]} {row[7]})" for row in pending_out]
                selected_out = st.selectbox(
                    "เลือกรายการที่ รพช. ส่งยามาคืนแล้ว:",
                    ["-- เลือกรายการ --"] + out_options,
                    key="refer_return_select"
                )
                
                if selected_out != "-- เลือกรายการ --":
                    doc_id = selected_out.split(" | ")[0].replace("เลขที่: ", "")
                    if st.button("✅ ยืนยันการรับยาคืนเข้าสต๊อก", key="confirm_refer_return"):
                        with st.spinner("กำลังอัปเดตฐานข้อมูล..."):
                            is_saved, msg = save_to_google_sheets("Outbound_Refer", action="update", doc_id=doc_id, new_status="รับคืนแล้ว")
                            if is_saved:
                                st.success(f"🎉 อัปเดตสถานะ {doc_id} เป็น 'รับคืนแล้ว' สำเร็จ!")
                                st.rerun() 
                            else:
                                st.error(f"❌ ผิดพลาด: {msg}")
            else:
                st.info("✨ ไม่มีรายการยารอรับคืนค่ะ")
        else:
            st.warning("กำลังโหลดข้อมูล หรือยังไม่มีข้อมูลในระบบ")

    # ---------------- แท็บที่ 2: ส่งยาคืน (เรายืมเขา) ----------------
    with tab2:
        st.markdown("**รายการที่เรายืมยามา และยังไม่ได้ส่งคืน (พิมพ์ใบคืนยา)**")
        in_data = get_from_google_sheets("Inbound_Shortage")
        
        if in_data and len(in_data) > 1:
            # ค้นหารายการที่สถานะเป็น "รอคืนยา"
            pending_in = [row for row in in_data[1:] if len(row) > 7 and row[7] == "รอคืนยา"]
            
            if pending_in:
                # 1. ดึงรายชื่อ รพ. ทั้งหมดที่มีค้างคืน (เพื่อจัดกลุ่มการส่งหนังสือ)
                hospitals = list(set([row[6] for row in pending_in]))
                selected_hosp = st.selectbox("1. เลือกโรงพยาบาลที่จะส่งยาคืน:", ["-- เลือกโรงพยาบาล --"] + hospitals)
                
                if selected_hosp != "-- เลือกโรงพยาบาล --":
                    # 2. กรองเฉพาะรายการของ รพ. ที่เลือก
                    hosp_items = [row for row in pending_in if row[6] == selected_hosp]
                    
                    # 3. ให้เลือกรายการยา (เลือกได้หลายรายการ)
                    item_options = [f"เลขที่: {row[0]} | ยา: {row[3]} ({row[4]} {row[5]})" for row in hosp_items]
                    selected_items = st.multiselect("2. เลือกรายการยาที่ต้องการส่งคืน (เลือกได้มากกว่า 1 รายการ):", item_options)
                    
                    if selected_items:
                        st.info(f"✨ จำนวนรายการที่เลือก: {len(selected_items)} รายการ")
                        
                        if st.button("✅ ยืนยันการส่งยาคืน และ สร้างหนังสือขอคืนยา"):
                            success_count = 0
                            drug_details_list = []
                            
                            with st.spinner("กำลังอัปเดตฐานข้อมูล และสร้างเอกสาร..."):
                                for idx, item_str in enumerate(selected_items):
                                    # แกะข้อมูลออกจากตัวเลือกที่กด
                                    doc_id = item_str.split(" | ")[0].replace("เลขที่: ", "")
                                    drug_info = item_str.split(" | ")[1].replace("ยา: ", "")
                                    
                                    # อัปเดตสถานะใน Sheet เป็น 'ส่งคืนแล้ว'
                                    is_saved, msg = save_to_google_sheets("Inbound_Shortage", action="update", doc_id=doc_id, new_status="ส่งคืนแล้ว")
                                    if is_saved:
                                        success_count += 1
                                        # จัดฟอร์แมตข้อความรายการยาที่จะไปใส่ใน Word
                                        drug_details_list.append(f"\t\t{idx+1}. {drug_info} (อ้างอิงใบยืม: {doc_id})")
                                    else:
                                        st.error(f"❌ ผิดพลาดในการอัปเดต {doc_id}: {msg}")
                                        
                            if success_count == len(selected_items):
                                st.success(f"🎉 อัปเดตสถานะเป็น 'ส่งคืนแล้ว' สำเร็จทั้ง {success_count} รายการ!")
                                
                                # สร้างเอกสาร Word อัตโนมัติรวมทุกรายการที่เลือก
                                drug_details_str = "\n".join(drug_details_list)
                                try:
                                    doc_return = DocxTemplate("template_return_official.docx")
                                    context_return = {
                                        'target_hospital': selected_hosp,
                                        'drug_details': drug_details_str
                                    }
                                    doc_return.render(context_return)
                                    bio_return = io.BytesIO()
                                    doc_return.save(bio_return)
                                    
                                    st.download_button(
                                        label="📥 โหลดหนังสือขอคืนยา (ตราครุฑ)",
                                        data=bio_return.getvalue(),
                                        file_name=f"Return_Official_{selected_hosp}_{datetime.datetime.now().strftime('%y%m%d')}.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    )
                                except Exception as e:
                                    st.error(f"⚠️ เกิดข้อผิดพลาดในการสร้างเอกสาร: {e}\n(กรุณาตรวจสอบว่ามีไฟล์ template_return_official.docx อยู่ในโฟลเดอร์เดียวกัน)")
            else:
                st.info("✨ ไม่มีรายการยารอส่งคืนค่ะ")
        else:
            st.warning("กำลังโหลดข้อมูล หรือยังไม่มีข้อมูลในระบบ")
            
    # ---------------- แท็บที่ 3: รับยาคืน (รพ.อื่นยืมเรา) ----------------
    with tab3:
        st.markdown("**รายการที่ รพ.อื่น ยืมยาไป และเรายังไม่ได้รับคืน**")
        lend_data = get_from_google_sheets("Outbound_Lend")
        
        if lend_data and len(lend_data) > 1:
            pending_lend = [row for row in lend_data[1:] if len(row) > 7 and row[7] == "รอรับคืน"]
            
            if pending_lend:
                lend_options = [f"เลขที่: {row[0]} | รพ.ที่ยืม: {row[3]} | ยา: {row[4]} ({row[5]} {row[6]})" for row in pending_lend]
                selected_lend = st.selectbox("เลือกรพ.ที่นำยามาคืนแล้ว:", ["-- เลือกรายการ --"] + lend_options)
                
                if selected_lend != "-- เลือกรายการ --":
                    doc_id_lend = selected_lend.split(" | ")[0].replace("เลขที่: ", "")
                    if st.button("✅ ยืนยันการรับยาคืนเข้าคลัง"):
                        with st.spinner("กำลังอัปเดตฐานข้อมูล..."):
                            is_saved, msg = save_to_google_sheets("Outbound_Lend", action="update", doc_id=doc_id_lend, new_status="รับคืนแล้ว")
                            if is_saved:
                                st.success(f"🎉 อัปเดตสถานะ {doc_id_lend} เป็น 'รับคืนแล้ว' สำเร็จ!")
                                st.rerun() 
                            else:
                                st.error(f"❌ ผิดพลาด: {msg}")
            else:
                st.info("✨ ไม่มีรายการยารอรับคืนจาก รพ.อื่น ค่ะ")
        else:
            st.warning("กำลังโหลดข้อมูล หรือยังไม่มีข้อมูลในระบบ")
