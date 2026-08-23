import streamlit as st
import streamlit.components.v1 as components
import datetime
import io
import requests
from docxtpl import DocxTemplate

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

# --- ฟังก์ชันติดต่อ Google Sheets ---
def get_from_google_sheets(sheet_name):
    try:
        response = requests.get(f"{WEBHOOK_URL}?sheet_name={sheet_name}")
        if response.status_code == 200 and response.json().get("status") == "success":
            return response.json().get("data")
        return []
    except:
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
        response = requests.post(WEBHOOK_URL, json=payload, allow_redirects=True)
        if response.status_code == 200 and "success" in response.text:
            return True, "Success"
        return False, response.text
    except Exception as e:
        return False, str(e)


# ==========================================
# 🖥️ เริ่มต้นการสร้างหน้าจอแอปพลิเคชัน
# ==========================================
st.set_page_config(page_title="ระบบจัดการยืม-คืนยา โรงพยาบาลศรีสังวรสุโขทัย", layout="centered", page_icon="SSW_Logo.jpg")

# --- โค้ดฝังฟอนต์ Sarabun, ซ่อนลายน้ำ และปรับแต่งสำหรับมือถือ ---
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"], [class*="st-"], h1, h2, h3, h4, h5, h6, span, label {
        font-family: 'Sarabun', sans-serif !important;
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

# นำรายชื่อมาใส่ใน Dropdown
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
# 🔐 ระบบ Security แบ่งสิทธิ์ User / Admin
# ==========================================
if user_role == "warehouse":
    st.sidebar.markdown("**🔐 ยืนยันตัวตน (Admin คลังยา)**")
    admin_password = st.sidebar.text_input("รหัสผ่าน Admin", type="password")
    
    # รหัสผ่านเริ่มต้นคือ 1234
    if admin_password != "1234":
        st.sidebar.error("❌ รหัสผ่านไม่ถูกต้อง (สิทธิ์การเข้าถึงถูกจำกัด)")
        st.stop()
    st.sidebar.success("✅ เข้าสู่ระบบระดับ Admin (คลังยา)")
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
if user_role == "warehouse":
    menu_options.append("5. 🔄 ติดตามสถานะคลังยา (Admin)")

# แสดงเมนูตามสิทธิ์
menu = st.sidebar.radio("เลือกกรณีที่ต้องการ", menu_options)


# ==========================================
# เมนูที่ 1: จ่ายยาออก (Refer Out)
# ==========================================
if menu == "1. จ่ายยาออก (Refer รพช.)":
    st.subheader("📤 กรณีที่ 1: จ่ายยาให้ผู้ป่วย Refer กลับ รพช.")
    
    col1, col2 = st.columns(2)
    with col1:
        pt_name = st.text_input("ชื่อ-นามสกุล ผู้ป่วย")
        hn = st.text_input("รหัสประจำตัวผู้ป่วย (HN)")
        target_hosp = st.selectbox("รพช. ปลายทาง", ["ทุ่งเสลี่ยม", "ศรีสัชนาลัย", "ศรีนคร", "สวรรคโลก", "อื่นๆ"])
    with col2:
        date_out = st.date_input("วันที่ทำรายการ", datetime.date.today())
        auto_doc_no = f"REF-{datetime.datetime.now().strftime('%y%m%d-%H%M')}"
        doc_no = st.text_input("เลขที่ใบยืม (สร้างอัตโนมัติ)", value=auto_doc_no, disabled=True)
    
    st.markdown("**รายการยา/เวชภัณฑ์ที่ให้ยืม**")
    
    # ปรับให้ช่องชื่อยา จำนวน และหน่วย อยู่ติดกัน
    col_d1, col_d2, col_d3 = st.columns([2, 1, 1])
    with col_d1:
        drug_name = st.text_input("ชื่อยา หรือ เวชภัณฑ์")
    with col_d2:
        qty = st.number_input("จำนวน", min_value=1)
    with col_d3:
        # เพิ่ม แกลลอน ในตัวเลือก
        unit_choice = st.selectbox("หน่วย", ["เม็ด", "ไวอัล", "แอมพูล", "ขวด", "หลอด", "กล่อง", "แกลลอน", "set", "ชิ้น", "อื่นๆ"])
        # ถ้าเลือกอื่นๆ จะมีช่องกรอกโผล่มาตรงนี้เลย
        unit = st.text_input("ระบุหน่วย...") if unit_choice == "อื่นๆ" else unit_choice
        
    st.markdown("**เหตุผลความจำเป็น**")
    reason_choice = st.radio("เลือกเหตุผล:", ["Refer Back", "ผู้ป่วยฉุกเฉิน / อุบัติเหตุ", "เหตุผลอื่นๆ ...."], horizontal=True, label_visibility="collapsed")
    note = st.text_input("โปรดระบุเหตุผลอื่นๆ ...") if reason_choice == "เหตุผลอื่นๆ ...." else reason_choice
        
    st.markdown("---")
    if st.button("💾 บันทึกข้อมูล และ สร้างใบให้ยืมยา"):
        # แทนที่ total_value เดิมด้วย "-" เพื่อรักษารูปแบบคอลัมน์ใน Google Sheets ไว้
        row_data = [doc_no, str(date_out), user_name, target_hosp, hn, drug_name, qty, unit, "-", note, "รอคืนยา"]
        
        with st.spinner('กำลังบันทึกข้อมูลลง Google Sheets...'):
            is_saved, debug_msg = save_to_google_sheets("Outbound_Refer", row_data=row_data, action="append")
        
        if is_saved:
            st.success("✅ บันทึกข้อมูลลง Google Sheets สำเร็จ!")
            try:
                doc_refer = DocxTemplate("template_refer_out.docx")
                context_refer = {
                    'target_hospital': target_hosp, 'pt_name': pt_name, 'hn': hn,
                    # เอาส่วนแสดงราคาออก เหลือแค่ ชื่อยา จำนวน หน่วย
                    'drug_details': f"{drug_name} จำนวน {qty} {unit}",
                    'user_name': user_name, 'thai_date': get_thai_date(date_out)
                }
                doc_refer.render(context_refer)
                bio_refer = io.BytesIO()
                doc_refer.save(bio_refer)
                
                st.download_button(
                    label="📥 โหลดใบให้ยืมยา (ส่งตัวผู้ป่วย)",
                    data=bio_refer.getvalue(),
                    file_name=f"ReferOut_{doc_no}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"⚠️ เกิดข้อผิดพลาดในการสร้างเอกสาร: {e}")
        else:
            st.error(f"❌ ไม่สามารถบันทึกข้อมูลได้ สาเหตุ: {debug_msg}")


# ==========================================
# เมนูที่ 2: ยืมยาเข้า (Borrow In)
# ==========================================
elif menu == "2. ยืมยาเข้า (ยา รพ. เราไม่พอ)":
    st.subheader("📥 กรณีที่ 2: ยืมยาจาก รพ. อื่น (ยาขาดคลัง)")
    
    col1, col2 = st.columns(2)
    with col1:
        auto_borrow_no = f"REQ-{datetime.datetime.now().strftime('%y%m%d-%H%M')}"
        borrow_no = st.text_input("เลขที่การยืม (สร้างอัตโนมัติ)", value=auto_borrow_no, disabled=True)
        source_hosp = st.text_input("รพ. ที่เราต้องการขอยืม")
    with col2:
        date_req = st.date_input("วันที่แจ้งเรื่อง", datetime.date.today())
    
    # จัดกลุ่มให้ ชื่อยา, จำนวน, และหน่วย เรียงต่อกัน
    st.markdown("**รายการยาที่ต้องการยืม**")
    col_d1, col_d2, col_d3 = st.columns([2, 1, 1])
    with col_d1:
        drug_missing = st.text_input("ชื่อยาที่ขาด/ต้องการยืม")
    with col_d2:
        borrow_qty = st.number_input("จำนวนที่ต้องการยืม", min_value=1)
    with col_d3:
        # เพิ่ม แกลลอน ในตัวเลือก
        unit_choice2 = st.selectbox("หน่วย (ยืมเข้า)", ["เม็ด", "ไวอัล", "แอมพูล", "ขวด", "หลอด", "กล่อง", "แกลลอน", "set", "ชิ้น", "อื่นๆ"])
        # ถ้าเลือกอื่นๆ ช่องกรอกจะปรากฏในคอลัมน์เดียวกัน
        unit2 = st.text_input("ระบุหน่วย... (ยืมเข้า)") if unit_choice2 == "อื่นๆ" else unit_choice2

    st.markdown("---")
    
    if st.button("💾 บันทึกข้อมูลการยืมยาเข้าคลัง"):
        row_data_in = [borrow_no, str(date_req), user_name, drug_missing, borrow_qty, unit2, source_hosp, "รอคืนยา", "-"]
        
        with st.spinner('กำลังบันทึกข้อมูลลง Google Sheets...'):
            is_saved, debug_msg = save_to_google_sheets("Inbound_Shortage", row_data=row_data_in, action="append")
        
        if is_saved:
            st.success("✅ บันทึกข้อมูลกรณีรพ. ยืมยาเข้า ลง Google Sheets สำเร็จ!")
        else:
            st.error(f"❌ ไม่สามารถบันทึกข้อมูลได้ สาเหตุ: {debug_msg}")

    st.markdown("---")
    st.subheader("🖨️ พิมพ์เอกสารขอยืมยาเข้า รพ.")
    col_memo, col_official = st.columns(2)

    with col_memo:
        st.info("🌙 สำหรับเภสัชกรอยู่เวร")
        if st.button("📄 พิมพ์บันทึกข้อความ"):
            try:
                doc_memo = DocxTemplate("template_memo.docx")
                context_memo = {
                    'target_hospital': source_hosp, 'drug_details': f"{drug_missing} จำนวน {borrow_qty} {unit2}",
                    'user_name': user_name, 'thai_date': get_thai_date(date_req)
                }
                doc_memo.render(context_memo)
                bio_memo = io.BytesIO()
                doc_memo.save(bio_memo)
                st.download_button("📥 โหลดบันทึกข้อความ", data=bio_memo.getvalue(), file_name=f"Memo_{borrow_no}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"⚠️ เกิดข้อผิดพลาด: {e}")

    with col_official:
        st.success("☀️ สำหรับคลังยา (งานสารบรรณ)")
        formal_reason = st.selectbox("เลือกเหตุผลในหนังสือ:", ["ยาขาดชั่วคราว จำเป็นต้องใช้เร่งด่วน", "ไม่มียาในบัญชีของโรงพยาบาล", "เป็นยา จ2 ไม่มียาในบัญชียา"])
        if st.button("🦅 พิมพ์หนังสือตราครุฑ"):
            try:
                doc_official = DocxTemplate("template_official.docx")
                context_official = {
                    'target_hospital': source_hosp, 'reason': formal_reason,
                    'drug_details': f"{drug_missing} จำนวน {borrow_qty} {unit2}", 'thai_date': get_thai_date(date_req)
                }
                doc_official.render(context_official)
                bio_official = io.BytesIO()
                doc_official.save(bio_official)
                st.download_button("📥 โหลดหนังสือตราครุฑ", data=bio_official.getvalue(), file_name=f"Official_{borrow_no}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"⚠️ เกิดข้อผิดพลาด: {e}")

# ==========================================
# เมนูที่ 3: ให้ รพ.อื่นยืมยา (Lend Out)
# ==========================================
elif menu == "3. ให้ รพ.อื่นยืมยา (ยา รพ.อื่นขาด)":
    st.subheader("📤 กรณีที่ 3: ให้ รพ.อื่นยืมยา (ยา รพ.อื่นขาดคลัง)")
    
    # --- แถวที่ 1: เลขที่ให้ยืม และ วันที่ ---
    col1, col2 = st.columns(2)
    with col1:
        auto_lend_no = f"LEND-{datetime.datetime.now().strftime('%y%m%d-%H%M')}"
        lend_no = st.text_input("เลขที่การให้ยืม (สร้างอัตโนมัติ)", value=auto_lend_no, disabled=True)
    with col2:
        date_lend = st.date_input("วันที่ให้ยืม", datetime.date.today())
        
    # --- แถวที่ 2: รพ. ที่มายืมยา ---
    target_hosp_lend = st.text_input("ชื่อ รพ. ที่มายืมยา")
    
    # --- แถวที่ 3: รายการยา จำนวน และหน่วย (เรียงติดกัน) ---
    st.markdown("**รายการยาที่ให้ยืม**")
    col_d1, col_d2, col_d3 = st.columns([2, 1, 1])
    
    with col_d1:
        drug_lended = st.text_input("ชื่อยา หรือ เวชภัณฑ์ที่ให้ยืม")
    with col_d2:
        lend_qty = st.number_input("จำนวนที่ให้ยืม", min_value=1, key="lend_qty")
    with col_d3:
        # เพิ่ม 'แกลลอน' เข้าไปในตัวเลือกด้วย เพื่อให้สอดคล้องกับเมนูอื่น
        unit_choice_lend = st.selectbox("หน่วย", ["เม็ด", "ไวอัล", "แอมพูล", "ขวด", "หลอด", "กล่อง", "แกลลอน", "set", "ชิ้น", "อื่นๆ"], key="unit_lend_select")
        unit_lend = st.text_input("ระบุหน่วย...") if unit_choice_lend == "อื่นๆ" else unit_choice_lend

    st.markdown("---")
    
    if st.button("💾 บันทึกข้อมูลให้ รพ.อื่นยืมยา"):
        row_data_lend = [lend_no, str(date_lend), user_name, target_hosp_lend, drug_lended, lend_qty, unit_lend, "รอรับคืน"]
        
        with st.spinner('กำลังบันทึกข้อมูลลง Google Sheets...'):
            is_saved, debug_msg = save_to_google_sheets("Outbound_Lend", row_data=row_data_lend, action="append")
        
        if is_saved:
            st.success(f"✅ บันทึกข้อมูลสำเร็จ! (รพ.{target_hosp_lend} ยืม {drug_lended} จำนวน {lend_qty} {unit_lend})")
        else:
            st.error(f"❌ ไม่สามารถบันทึกข้อมูลได้ สาเหตุ: {debug_msg}")

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
    
    tab1, tab2, tab3 = st.tabs(["📥 รับยาคืน (Refer Out)", "📤 ส่งยาคืน (เรายืมเขา)", "📥 รับยาคืน (รพ.อื่นยืมเรา)"])
    
    # ---------------- แท็บที่ 1: รับยาคืน ----------------
    with tab1:
        st.markdown("**รายการที่จ่ายยาออกไป และยังไม่ได้รับคืน**")
        out_data = get_from_google_sheets("Outbound_Refer")
        
        if out_data and len(out_data) > 1:
            pending_out = [row for row in out_data[1:] if len(row) > 10 and row[10] == "รอคืนยา"]
            
            if pending_out:
                out_options = [f"เลขที่: {row[0]} | รพ: {row[3]} | ยา: {row[5]} ({row[6]} {row[7]})" for row in pending_out]
                selected_out = st.selectbox("เลือกรายการที่ รพช. ส่งยามาคืนแล้ว:", ["-- เลือกรายการ --"] + out_options)
                
                if selected_out != "-- เลือกรายการ --":
                    doc_id = selected_out.split(" | ")[0].replace("เลขที่: ", "")
                    if st.button("✅ ยืนยันการรับยาคืนเข้าสต๊อก"):
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
