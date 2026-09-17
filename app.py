import streamlit as st
import os
import json
import io
import copy
import time
from datetime import datetime, timedelta
from docxtpl import DocxTemplate
import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from google import genai
from google.genai import types

# กำหนดรหัสผ่านสำหรับปลดล็อกระบบ
SYSTEM_PASSCODE = "0863449483"

st.set_page_config(
    page_title="ระบบบันทึกหลังการสอน AI อาชีวศึกษา",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# สไตล์ตกแต่ง UI สีสัน สดใส มีมิติ พร้อมกล่องลิขสิทธิ์
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Prompt', sans-serif;
    }
    .main-header {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 50%, #06B6D4 100%);
        padding: 24px;
        border-radius: 16px;
        color: white;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px rgba(30, 58, 138, 0.25);
    }
    .main-header h1 {
        color: white !important;
        font-weight: 700;
        margin-bottom: 8px;
        text-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }
    .main-header p {
        color: #E0F2FE !important;
        font-size: 15px;
        margin-bottom: 0;
    }
    .card-box {
        background: #ffffff;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
    }
    .badge-tag {
        background: linear-gradient(90deg, #EC4899, #8B5CF6);
        color: white;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
        display: inline-block;
        margin-bottom: 10px;
    }
    .stButton>button {
        background: linear-gradient(90deg, #F43F5E 0%, #E11D48 100%) !important;
        color: white !important;
        font-size: 17px !important;
        font-weight: 600 !important;
        border-radius: 12px !important;
        border: none !important;
        padding: 12px 24px !important;
        box-shadow: 0 8px 20px rgba(225, 29, 72, 0.3) !important;
        transition: all 0.3s ease !important;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 12px 25px rgba(225, 29, 72, 0.45) !important;
    }
    .footer-box {
        text-align: center;
        padding: 24px 10px;
        margin-top: 50px;
        border-top: 1px solid #E2E8F0;
        color: #64748B;
        font-size: 13.5px;
    }
    .footer-badge {
        display: inline-block;
        background: #F1F5F9;
        border: 1px solid #CBD5E1;
        padding: 5px 14px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
        color: #334155;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ระบบตรวจสอบรหัสผ่านปลดล็อก
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
    <div class="main-header">
        <h1>🔒 ระบบจัดทำบันทึกหลังการสอนอัตโนมัติ AI</h1>
        <p>กรุณากรอกรหัสผ่านเพื่อปลดล็อกเข้าสู่ระบบ</p>
    </div>
    """, unsafe_allow_html=True)

    col_l1, col_l2, col_l3 = st.columns([1, 1.2, 1])
    with col_l2:
        st.markdown('<div class="card-box">', unsafe_allow_html=True)
        st.subheader("🔑 ยืนยันสิทธิ์การเข้าใช้งาน")
        pass_input = st.text_input("รหัสปลดล็อกระบบ (Passcode):", type="password", placeholder="กรอกรหัสจากผู้ดูแลระบบ...")
        if st.button("🔓 ปลดล็อกเข้าสู่ระบบ", use_container_width=True):
            if pass_input == SYSTEM_PASSCODE:
                st.session_state.authenticated = True
                st.success("✅ ปลดล็อกสำเร็จ กำลังเข้าสู่ระบบ...")
                st.rerun()
            else:
                st.error("❌ รหัสผ่านไม่ถูกต้อง กรุณาติดต่อเจ้าของระบบ")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown("""
        <div class="footer-box" style="margin-top: 20px;">
            <div class="footer-badge">🛡️ PROPRIETARY SOFTWARE</div><br/>
            © สงวนลิขสิทธิ์ พัฒนาโดย <b>นายณัฐวุฒิ หล้าปงสาย</b> ครูผู้ช่วย วิทยาลัยเทคนิคจันทบุรี
        </div>
        """, unsafe_allow_html=True)
    st.stop()

# เมื่อปลดล็อกผ่าน เข้าสู่หน้าหลักของระบบ
THAI_MONTHS = [
    "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]
DAY_NAMES = ["วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์", "วันอาทิตย์"]
DEPARTMENT_OPTIONS = [
    "การจัดการโลจิสติกส์และซัพพลายเชน",
    "เทคโนโลยีสารสนเทศ",
    "คอมพิวเตอร์ธุรกิจ",
    "การบัญชี",
    "การตลาด",
    "ช่างยนต์",
    "ช่างไฟฟ้ากำลัง",
    "ช่างอิเล็กทรอนิกส์",
    "ช่างก่อสร้าง",
    "อื่นๆ (ระบุเอง)"
]

st.markdown("""
<div class="main-header">
    <h1>📝 ระบบจัดทำบันทึกหลังการสอนอัตโนมัติ (AI Professional)</h1>
    <p>วิเคราะห์โครงการสอน สกัดรายสัปดาห์ รองรับการฉีกคาบสูงสุด 4 คาบ และเรนเดอร์ลงแบบฟอร์มวิทยาลัยเป๊ะ 100%</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ การตั้งค่าระบบ & สิทธิ์")
    st.markdown('<span class="badge-tag">STATUS: UNLOCKED</span>', unsafe_allow_html=True)
    
    api_key_input = st.text_input("🔑 Gemini API Key:", type="password", placeholder="AIzaSy...")
    st.markdown("[👉 รับ API Key ฟรีคลิกที่นี่](https://aistudio.google.com/apikey)")
    st.divider()

    st.subheader("👤 ข้อมูลครูผู้สอน")
    teacher_name = st.text_input("ชื่อ-สกุลครูผู้สอน:", value="นายณัฐวุฒิ หล้าปงสาย")
    
    dept_choice = st.selectbox("สาขาวิชา / แผนกวิชา:", DEPARTMENT_OPTIONS, index=0)
    if dept_choice == "อื่นๆ (ระบุเอง)":
        department = st.text_input("ระบุสาขาวิชาของคุณ:", value="")
    else:
        department = dept_choice

    st.markdown("---")
    if st.button("🔒 ล็อกระบบกลับ"):
        st.session_state.authenticated = False
        st.rerun()

    st.markdown("""
    <div style="font-size: 12px; color: #94A3B8; text-align: center; margin-top: 25px;">
        <b>AI Vocational Reflection System</b><br/>
        สงวนลิขสิทธิ์ พัฒนาโดย<br/>
        <b>นายณัฐวุฒิ หล้าปงสาย</b><br/>
        ครูผู้ช่วย วิทยาลัยเทคนิคจันทบุรี
    </div>
    """, unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 1. แบบฟอร์มและหลักสูตร")
    tpl_file = st.file_uploader("📄 แนบแบบฟอร์มวิทยาลัย (template.docx):", type=["docx"])
    uploaded_file = st.file_uploader("📚 แนบไฟล์โครงการสอน (PDF, Word, TXT, รูปภาพ):", type=["pdf", "docx", "txt", "png", "jpg", "jpeg"])
    
    st.markdown("---")
    st.markdown("🎯 **เลือกระดับชั้นและวุฒิการศึกษา:**")
    c_deg, c_yr = st.columns(2)
    with c_deg:
        degree = st.selectbox("ระดับคุณวุฒิ:", ["ปวช.", "ปวส."])
    with c_yr:
        if degree == "ปวช.":
            year_num = st.selectbox("ชั้นปี:", ["1", "2", "3"])
            target_weeks = 18
        else:
            year_num = st.selectbox("ชั้นปี:", ["1", "2"])
            target_weeks = 15

    class_level = f"{degree} {year_num}"
    st.info(f"✨ ระดับ: **{class_level}** | สาขา: **{department}** | กำหนดอัตโนมัติ: **{target_weeks} สัปดาห์**")

with col2:
    st.subheader("⏰ 2. ตารางวัน-เวลา และการฉีกคาบสอน")
    slots_count = st.selectbox(
        "จำนวนคาบสอนใน 1 สัปดาห์ (ฉีกคาบได้สูงสุด 4 คาบ):",
        options=[1, 2, 3, 4],
        format_func=lambda x: f"สอน {x} คาบ / สัปดาห์" if x > 1 else "สอน 1 คาบ (วันเดียวจบ)",
        index=1
    )

    slots_info = []
    default_days = [0, 1, 2, 3]
    default_times = ["15.30-16.30 น.", "08.30-10.30 น.", "10.30-12.30 น.", "13.30-15.30 น."]

    for i in range(slots_count):
        st.markdown(f"**📌 รายละเอียดคาบที่ {i+1}:**")
        sc1, sc2 = st.columns(2)
        with sc1:
            d_val = st.selectbox(f"วัน (คาบที่ {i+1}):", DAY_NAMES, index=default_days[i % len(default_days)], key=f"day_slot_{i}")
        with sc2:
            t_val = st.text_input(f"เวลา (คาบที่ {i+1}):", value=default_times[i % len(default_times)], key=f"time_slot_{i}")
        slots_info.append({"day": d_val, "time": t_val})

    start_date = st.date_input("📅 วันที่เริ่มสอนสัปดาห์ที่ 1 (คำนวณปฏิทินไทยอัตโนมัติ):")

    holiday_text = st.text_area(
        "ระบุสัปดาห์และเหตุผลการงดสอน (ถ้ามี เช่น '8:ตรงกับวันหยุดนักขัตฤกษ์'):",
        value="8:ตรงกับวันหยุดนักขัตฤกษ์ตามประกาศสถานศึกษา",
        height=70
    )

def format_thai_date(dt):
    d = dt.day
    m = THAI_MONTHS[dt.month]
    y = dt.year + 543
    day_name = DAY_NAMES[dt.weekday()]
    return f"{day_name} {d} {m} {y}"

if st.button(f"🚀 เริ่มสร้างเอกสารบันทึกหลังการสอนครบ {target_weeks} สัปดาห์", use_container_width=True):
    api_key = api_key_input.strip() if api_key_input else ""
    if not api_key:
        st.warning("⚠️ กรุณากรอก Gemini API Key ที่แถบด้านซ้ายก่อนเริ่มใช้งาน")
        st.stop()
    if not tpl_file:
        st.warning("⚠️ กรุณาแนบไฟล์ template.docx ของวิทยาลัย")
        st.stop()
    if not uploaded_file:
        st.warning("⚠️ กรุณาแนบไฟล์โครงการสอน")
        st.stop()
    if not department.strip():
        st.warning("⚠️ กรุณาระบุสาขาวิชา/แผนกวิชา")
        st.stop()

    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        status_text.text("🤖 กำลังส่งข้อมูลให้ Gemini AI วิเคราะห์โครงการสอน...")
        client = genai.Client(api_key=api_key)
        file_bytes = uploaded_file.read()
        mime_type = uploaded_file.type or "application/octet-stream"

        prompt = f"""
        วิเคราะห์เนื้อหาโครงการสอนที่แนบมานี้ และจัดทำเนื้อหาบันทึกหลังการสอนอาชีวศึกษา
        สำหรับระดับชั้น {class_level} ให้ครบถ้วนตั้งแต่สัปดาห์ที่ 1 ถึงสัปดาห์ที่ {target_weeks} (รวม {target_weeks} สัปดาห์พอดี ห้ามขาด)
        ข้อมูลวันหยุด/งดสอน: {holiday_text}

        ข้อกำหนดสำคัญเพื่อไม่ให้หน้ากระดาษล้น:
        1. topic: สรุปสั้นๆ 1 บรรทัด
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
            "gemini-3.8-flash",
            "gemini-3.5-flash-lite"
        ]
        response = None
        last_error = None

        for target_m in models_to_try:
            status_text.text(f"⏳ กำลังประมวลผลด้วยโมเดล {target_m}...")
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
        day_map = {name: idx for idx, name in enumerate(DAY_NAMES)}

        for idx, w in enumerate(final_weeks):
            progress_bar.progress(int(((idx + 1) / total_count) * 100))
            status_text.text(f"📝 กำลังลงข้อมูลสัปดาห์ที่ {w.get('week')} ในแบบฟอร์มวิทยาลัย...")

            week_num = w.get("week", idx + 1)
            base_week_date = start_date + timedelta(weeks=(week_num - 1))
            
            date_lines = []
            time_lines = []
            for slot in slots_info:
                t_wday = day_map.get(slot["day"], 0)
                dt_slot = base_week_date + timedelta(days=(t_wday - base_week_date.weekday()))
                date_lines.append(format_thai_date(dt_slot))
                time_lines.append(f"เวลา {slot['time']}")

            date_display = "\n".join(date_lines)
            time_display = "\n".join(time_lines)

            is_hol = w.get("is_holiday", False)
            context_w = {
                "code": course_code,
                "subject": course_name,
                "week": week_num,
                "date": date_display,
                "date_display": date_display,
                "time": time_display,
                "time_display": time_display,
                "topic": w.get("topic"),
                "level": class_level,
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
                # รวมเอกสารโดยต่อหน้าใหม่แบบไม่มีหน้าว่างคั่น
                first_element = True
                for el in sub_doc.element.body:
                    if el.tag.endswith('sectPr'):
                        continue
                    copied_el = copy.deepcopy(el)
                    if first_element:
                        # บังคับให้องค์ประกอบแรกขึ้นหน้าใหม่ทันทีโดยไม่สร้างย่อหน้าเปล่า
                        if copied_el.tag.endswith('p'):
                            pPr = copied_el.find('{[http://schemas.openxmlformats.org/wordprocessingml/2006/main](http://schemas.openxmlformats.org/wordprocessingml/2006/main)}pPr')
                            if pPr is None:
                                pPr = parse_xml(r'<w:pPr %s><w:pageBreakBefore/></w:pPr>' % nsdecls('w'))
                                copied_el.insert(0, pPr)
                            else:
                                pPr.append(parse_xml(r'<w:pageBreakBefore %s/>' % nsdecls('w')))
                        elif copied_el.tag.endswith('tbl'):
                            # ถ้าเป็นตาราง ให้แทรกตัวแบ่งหน้าบน Paragraph นำหน้าแบบแนบชิด
                            p_break = parse_xml(r'<w:p %s><w:pPr><w:pageBreakBefore/><w:spacing w:after="0" w:before="0" w:line="1" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>' % nsdecls('w'))
                            merged_doc.element.body.append(p_break)
                        first_element = False
                    merged_doc.element.body.append(copied_el)

        output_stream = io.BytesIO()
        merged_doc.save(output_stream)
        output_stream.seek(0)

        progress_bar.progress(100)
        status_text.empty()

        st.balloons()
        st.success(f"🎉 สร้างเอกสารสำเร็จครบ {target_weeks} สัปดาห์ เรียงต่อกันหน้าต่อหน้า ไม่มีหน้าว่าง 100%!")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Word (แบบฟอร์มวิทยาลัยตรงเป๊ะ)",
            data=output_stream,
            file_name=f"บันทึกหลังการสอน_{course_code}_{class_level.replace(' ', '')}_ครบ{target_weeks}สัปดาห์.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    except Exception as e:
        status_text.empty()
        st.error(f"เกิดข้อผิดพลาด: {str(e)}")

# กล่องข้อมูลลิขสิทธิ์และผู้พัฒนาระบบด้านล่างสุด
st.markdown("""
<div class="footer-box">
    <div class="footer-badge">🛡️ PROPRIETARY & EDUCATIONAL OPEN-SOURCE</div><br/>
    <b>ระบบปัญญาประดิษฐ์สกัดและจัดทำบันทึกหลังการสอนอาชีวศึกษา (AI Vocational Reflection)</b><br/>
    สงวนลิขสิทธิ์ พัฒนาโดย <b>นายณัฐวุฒิ หล้าปงสาย</b> ครูผู้ช่วย วิทยาลัยเทคนิคจันทบุรี<br/>
    <span style="font-size: 12px; color: #94A3B8;">ขับเคลื่อนด้วย Streamlit & Google Gemini AI Flash Engine</span>
</div>
""", unsafe_allow_html=True)
