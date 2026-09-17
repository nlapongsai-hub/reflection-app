import streamlit as st
import os
import json
import io
import copy
import time
from pathlib import Path
from docxtpl import DocxTemplate
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from google import genai
from google.genai import types

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="ระบบบันทึกหลังการสอน AI | อาชีวศึกษา",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# CUSTOM CSS: สไตล์ Dashboard แบบ EdTech โมเดิร์น
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Sarabun', sans-serif !important;
    }

    /* พื้นหลังหลักของเว็บ */
    .stApp {
        background-color: #f8fafc;
    }

    /* Header Bar ด้านบน */
    .top-navbar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: #ffffff;
        padding: 12px 24px;
        border-radius: 16px;
        box-shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.05);
        border: 1px solid #e2e8f0;
        margin-bottom: 25px;
    }
    .brand-title {
        font-size: 20px;
        font-weight: 700;
        color: #1e293b;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .status-badge {
        background-color: #ecfdf5;
        color: #059669;
        border: 1px solid #a7f3d0;
        padding: 6px 14px;
        border-radius: 9999px;
        font-size: 13px;
        font-weight: 600;
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }
    .user-badge {
        display: flex;
        align-items: center;
        gap: 10px;
        background: #f1f5f9;
        padding: 6px 14px;
        border-radius: 9999px;
        font-size: 13px;
        color: #334155;
        font-weight: 500;
    }

    /* กล่องการ์ดเนื้อหา */
    .dashboard-card {
        background: #ffffff;
        border-radius: 18px;
        padding: 24px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 25px -4px rgba(0, 0, 0, 0.04);
        margin-bottom: 20px;
    }
    .card-title {
        font-size: 17px;
        font-weight: 700;
        color: #1e3a8a;
        margin-bottom: 15px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    /* ปรับแต่งปุ่ม Action Button */
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 16px !important;
        padding: 14px 24px !important;
        border-radius: 12px !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.35) !important;
        transition: all 0.2s ease-in-out !important;
    }
    div.stButton > button:first-child:hover {
        background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%) !important;
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.45) !important;
        transform: translateY(-1px) !important;
    }

    /* ปรับแต่งปุ่มดาวน์โหลด */
    div.stDownloadButton > button {
        background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 16px !important;
        padding: 14px 24px !important;
        border-radius: 12px !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(5, 150, 105, 0.35) !important;
    }

    /* ซ่อนแถบ Header ส่วนเกินของ Streamlit */
    header[data-testid="stHeader"] {
        background-color: transparent !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# TOP BAR / HEADER
# ==========================================
st.markdown("""
<div class="top-navbar">
    <div class="brand-title">
        <span>📘</span> ระบบจัดทำบันทึกหลังการสอนอัตโนมัติ (AI Teacher Assistant)
    </div>
    <div style="display: flex; gap: 12px; align-items: center;">
        <span class="status-badge">🟢 Google Gemini: พร้อมใช้งาน</span>
        <span class="user-badge">👤 นายณัฐวุฒิ หล้าปงสาย (วิทยาลัยเทคนิคจันทบุรี)</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# SIDEBAR: การตั้งค่าระบบ
# ==========================================
with st.sidebar:
    st.markdown("### ⚙️ การตั้งค่าระบบ AI")
    api_key_input = st.text_input("🔑 Gemini API Key:", type="password", placeholder="AIzaSy...")
    st.caption("[คลิกที่นี่เพื่อขอรับ API Key ฟรีจาก Google AI Studio](https://aistudio.google.com/apikey)")
    
    st.markdown("---")
    st.markdown("### 👨‍🏫 ข้อมูลผู้สอนและรายวิชา")
    teacher_name = st.text_input("ชื่อ-สกุลครูผู้สอน:", value="นายณัฐวุฒิ หล้าปงสาย")
    department = st.text_input("แผนกวิชา:", value="การจัดการโลจิสติกส์และซัพพลายเชน")
    time_slot = st.text_input("ช่วงเวลาเรียน:", value="08.30 - 10.30 น.")

# ==========================================
# MAIN INTERFACE: การ์ดทำงาน
# ==========================================
col1, col2 = st.columns([1.1, 0.9], gap="large")

with col1:
    st.markdown("""
    <div class="dashboard-card">
        <div class="card-title">📂 แนบไฟล์ข้อมูลโครงการสอน & แบบฟอร์ม</div>
    """, unsafe_allow_html=True)
    
    tpl_file = st.file_uploader(
        "1. แนบแบบฟอร์มวิทยาลัย (template.docx):",
        type=["docx"],
        help="ไฟล์ Word ที่มีตราวิทยาลัยและตารางแม่แบบ"
    )
    uploaded_file = st.file_uploader(
        "2. แนบไฟล์โครงการสอน (PDF, DOCX, รูปภาพ):",
        type=["pdf", "docx", "txt", "png", "jpg", "jpeg"],
        help="โครงการสอนรายวิชาตลอดภาคเรียน"
    )
    st.markdown("</div>", unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="dashboard-card">
        <div class="card-title">🎯 เงื่อนไขการจัดทำบันทึก</div>
    """, unsafe_allow_html=True)
    
    level_type = st.radio("ระดับการศึกษา / จำนวนสัปดาห์:", ["ปวช. (18 สัปดาห์)", "ปวส. (15 สัปดาห์)"], horizontal=True)
    target_weeks = 18 if "18" in level_type else 15
    
    holiday_text = st.text_area(
        "ระบุสัปดาห์และเหตุผลการงดสอน (ถ้ามี):",
        value="8:ตรงกับวันหยุดนักขัตฤกษ์ตามประกาศสถานศึกษา",
        height=110,
        help="รูปแบบ: สัปดาห์:เหตุผล (แยกบรรทัดได้)"
    )
    st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# ปุ่มประมวลผล
# ==========================================
if st.button("🚀 เริ่มวิเคราะห์และสร้างเอกสารบันทึกหลังการสอนครบทุกสัปดาห์", type="primary", use_container_width=True):
    api_key = api_key_input.strip() if api_key_input else ""
    if not api_key:
        st.warning("⚠️ กรุณากรอก Gemini API Key ที่แถบด้านซ้ายก่อนเริ่มใช้งาน")
        st.stop()
    if not tpl_file:
        st.warning("⚠️ กรุณาแนบไฟล์ template.docx")
        st.stop()
    if not uploaded_file:
        st.warning("⚠️ กรุณาแนบไฟล์โครงการสอน")
        st.stop()

    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        status_text.text("กำลังส่งข้อมูลให้ Gemini AI วิเคราะห์โครงการสอน...")
        client = genai.Client(api_key=api_key)
        file_bytes = uploaded_file.read()
        mime_type = uploaded_file.type or "application/octet-stream"

        prompt = f"""
        วิเคราะห์เนื้อหาโครงการสอนที่แนบมานี้ และจัดทำเนื้อหาบันทึกหลังการสอนอาชีวศึกษา
        สำหรับระดับชั้น {level_type} ให้ครบถ้วนตั้งแต่สัปดาห์ที่ 1 ถึงสัปดาห์ที่ {target_weeks} (รวม {target_weeks} สัปดาห์พอดี ห้ามขาด)
        ช่วงเวลาสอน: {time_slot}
        ข้อมูลวันหยุด/งดสอน: {holiday_text}

        ข้อกำหนดสำคัญเพื่อความกระชับและไม่ให้หน้ากระดาษล้น:
        1. topic: เขียนให้กระชับ ชัดเจน ไม่เกิน 1 บรรทัด
        2. student_eval: สรุปผลด้าน K, P, A และร้อยละผู้เรียนที่ผ่านเกณฑ์ ความยาว 1-2 บรรทัด
        3. teacher_eval: สรุปกิจกรรมและสื่อที่ใช้ ความยาว 1-2 บรรทัด
        4. problem_solution: สรุปปัญหาและวิธีแก้ไข ความยาว 1-2 บรรทัด

        ส่งออกเป็น Pure JSON โครงสร้างนี้เท่านั้น:
        {{
            "code": "รหัสวิชา",
            "subject": "ชื่อวิชา",
            "weeks": [
                {{
                    "week": 1,
                    "date": "สัปดาห์ที่ 1",
                    "time": "{time_slot}",
                    "topic": "ชื่อหน่วยและเรื่องที่สอน",
                    "is_holiday": false,
                    "off_reason": "-",
                    "student_eval": "ผลด้านผู้เรียน",
                    "teacher_eval": "ผลด้านผู้สอน",
                    "problem_solution": "ปัญหาและแนวทางแก้ไข"
                }}
            ]
        }}
        ห้ามใส่เครื่องหมาย markdown block ส่งเฉพาะ Pure JSON
        """

        models_to_try = [
            "gemini-3.5-flash-lite",
            "gemini-3.6-flash"
        ]

        response = None
        last_error = None

        for target_m in models_to_try:
            status_text.text(f"กำลังประมวลผลด้วยโมเดล {target_m}...")
            try:
                response = client.models.generate_content(
                    model=target_m,
                    contents=[
                        types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
                        prompt
                    ],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.2
                    )
                )
                if response and response.text:
                    break
            except Exception as err:
                last_error = err
                time.sleep(1)

        if not response or not response.text:
            raise Exception(f"ไม่สามารถเชื่อมต่อโมเดลได้: {last_error}")

        clean_text = response.text.strip()
        if clean_text.startswith("```json"):
            clean_text = clean_text[7:]
        if clean_text.startswith("```"):
            clean_text = clean_text[3:]
        if clean_text.endswith("```"):
            clean_text = clean_text[:-3]

        data = json.loads(clean_text)
        course_code = data.get("code", "วิชา")
        course_name = data.get("subject", "โครงการสอน")
        weeks_data = data.get("weeks", [])

        existing = {w.get("week"): w for w in weeks_data}
        final_weeks = []
        for i in range(1, target_weeks + 1):
            if i in existing:
                final_weeks.append(existing[i])
            else:
                final_weeks.append({
                    "week": i,
                    "date": f"สัปดาห์ที่ {i}",
                    "time": time_slot,
                    "topic": f"หน่วยการเรียนรู้ที่ {i}",
                    "is_holiday": False,
                    "off_reason": "-",
                    "student_eval": "ผู้เรียนผ่านเกณฑ์ร้อยละ 90 ขึ้นไป มีทักษะและความตั้งใจในการปฏิบัติงาน",
                    "teacher_eval": "จัดการเรียนรู้เชิงรุก (Active Learning) ผู้เรียนมีส่วนร่วมได้ดี",
                    "problem_solution": "ให้คำแนะนำเพิ่มเติมแก่นักเรียนรายบุคคลหลังเลิกเรียน"
                })

        tpl_bytes = tpl_file.read()
        merged_doc = None
        total_count = len(final_weeks)
        cur_level = "ปวช." if target_weeks == 18 else "ปวส."

        for idx, w in enumerate(final_weeks):
            progress_bar.progress(int(((idx + 1) / total_count) * 100))
            status_text.text(f"กำลังลงข้อมูลสัปดาห์ที่ {w.get('week')} ในแบบฟอร์มวิทยาลัย...")

            is_hol = w.get("is_holiday", False)
            context_w = {
                "code": course_code,
                "subject": course_name,
                "week": w.get("week"),
                "date": w.get("date"),
                "time": w.get("time", time_slot),
                "topic": w.get("topic"),
                "level": cur_level,
                "department": department,
                "check_on": "☑" if not is_hol else "☐",
                "check_off": "☑" if is_hol else "☐",
                "off_reason": w.get("off_reason", "-") if is_hol else "-",
                "student_eval": w.get("student_eval"),
                "teacher_eval": w.get("teacher_eval"),
                "problem_solution": w.get("problem_solution"),
            }

            t = DocxTemplate(io.BytesIO(tpl_bytes))
            render_context = {
                "teacher_name": teacher_name,
                "w": context_w,
                **context_w
            }
            t.render(render_context)

            tmp_io = io.BytesIO()
            t.save(tmp_io)
            tmp_io.seek(0)

            sub_doc = docx.Document(tmp_io)
            
            if merged_doc is None:
                merged_doc = sub_doc
            else:
                first_element = True
                for el in sub_doc.element.body:
                    if el.tag.endswith('sectPr'):
                        continue
                    
                    new_el = copy.deepcopy(el)
                    
                    if first_element:
                        if new_el.tag.endswith('tbl'):
                            p_break = OxmlElement('w:p')
                            pPr = OxmlElement('w:pPr')
                            pPr.append(OxmlElement('w:pageBreakBefore'))
                            p_break.append(pPr)
                            merged_doc.element.body.append(p_break)
                        elif new_el.tag.endswith('p'):
                            pPr = new_el.find(qn('w:pPr'))
                            if pPr is None:
                                pPr = OxmlElement('w:pPr')
                                new_el.insert(0, pPr)
                            pPr.append(OxmlElement('w:pageBreakBefore'))
                        first_element = False
                    
                    merged_doc.element.body.append(new_el)

        output_stream = io.BytesIO()
        merged_doc.save(output_stream)
        output_stream.seek(0)

        progress_bar.progress(100)
        status_text.empty()

        st.success(f"🎉 สร้างเอกสารสำเร็จครบ {target_weeks} สัปดาห์! ({course_code} {course_name})")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Word (แบบฟอร์มวิทยาลัยเป๊ะ ไร้หน้าว่าง)",
            data=output_stream,
            file_name=f"บันทึกหลังการสอน_{course_code}_ครบ{target_weeks}สัปดาห์.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    except Exception as e:
        status_text.empty()
        st.error(f"เกิดข้อผิดพลาด: {str(e)}")

# ==========================================
# FOOTER: ลิขสิทธิ์และเครดิตผู้พัฒนา
# ==========================================
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: #64748b; font-size: 14.5px; line-height: 1.8; margin-top: 15px; margin-bottom: 25px;">
        พัฒนาโดย <strong style="color: #334155;">นายณัฐวุฒิ หล้าปงสาย</strong><br>
        ครูผู้ช่วย วิทยาลัยเทคนิคจันทบุรี
    </div>
    """,
    unsafe_allow_html=True
)
