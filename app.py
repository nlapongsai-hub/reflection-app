import streamlit as st
import os
import json
import io
import copy
import time
from datetime import datetime, timedelta
from docxtpl import DocxTemplate
import docx
from google import genai
from google.genai import types

st.set_page_config(page_title="ระบบบันทึกหลังการสอน AI", page_icon="📝", layout="wide")

st.title("📝 ระบบจัดทำบันทึกหลังการสอนอัตโนมัติ")
st.caption("สกัดข้อมูลจากโครงการสอนและสร้างเอกสาร Word ตามแบบฟอร์มวิทยาลัยเป๊ะ 100%")

THAI_MONTHS = [
    "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]
DAY_NAMES = ["วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์", "วันอาทิตย์"]

with st.sidebar:
    st.header("⚙️ การตั้งค่าระบบ")
    api_key_input = st.text_input("Gemini API Key:", type="password", placeholder="AIzaSy...")
    st.markdown("[รับ API Key ฟรีที่นี่](https://aistudio.google.com/apikey)")
    st.divider()
    teacher_name = st.text_input("ชื่อ-สกุลครูผู้สอน:", value="นายณัฐวุฒิ ละผ่องใส")
    department = st.text_input("สาขาวิชา/แผนกวิชา:", value="การจัดการโลจิสติกส์และซัพพลายเชน")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. แบบฟอร์มและระดับการศึกษา")
    tpl_file = st.file_uploader("📄 แนบแบบฟอร์มวิทยาลัย (template.docx):", type=["docx"])
    uploaded_file = st.file_uploader("📚 แนบไฟล์โครงการสอน (PDF, Word, TXT, รูปภาพ):", type=["pdf", "docx", "txt", "png", "jpg", "jpeg"])
    
    st.markdown("---")
    st.markdown("**เลือกระดับชั้นการศึกษา:**")
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
    st.info(f"📌 ระดับชั้น: **{class_level}** | กำหนดจำนวนอัตโนมัติ: **{target_weeks} สัปดาห์**")

with col2:
    st.subheader("2. กำหนดตารางเวลาและวันสอน")
    schedule_mode = st.radio(
        "รูปแบบคาบสอนในแต่ละสัปดาห์:",
        ["สอนวันเดียว (รวดเดียว)", "ฉีกคาบสอน (แยก 2 วันใน 1 สัปดาห์)"],
        index=1
    )

    if schedule_mode == "สอนวันเดียว (รวดเดียว)":
        day_1 = st.selectbox("วันที่สอน:", DAY_NAMES, index=0)
        time_1 = st.text_input("ช่วงเวลาเรียน:", value="08.30 - 10.30 น.")
        is_split = False
    else:
        is_split = True
        c_sub1, c_sub2 = st.columns(2)
        with c_sub1:
            day_1 = st.selectbox("คาบที่ 1 (วัน):", DAY_NAMES, index=0)
            time_1 = st.text_input("เวลาคาบที่ 1:", value="15.30-16.30 น.")
        with c_sub2:
            day_2 = st.selectbox("คาบที่ 2 (วัน):", DAY_NAMES, index=1)
            time_2 = st.text_input("เวลาคาบที่ 2:", value="08.30-10.30 น.")

    start_date = st.date_input("วันที่เริ่มสอนสัปดาห์ที่ 1 (เพื่อคำนวณปฏิทินอัตโนมัติ):")

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

if st.button(f"🚀 เริ่มสร้างเอกสารบันทึกหลังการสอนครบ {target_weeks} สัปดาห์", type="primary", use_container_width=True):
    api_key = api_key_input.strip() if api_key_input else ""
    if not api_key:
        st.warning("กรุณากรอก Gemini API Key ที่แถบด้านซ้ายก่อนเริ่มใช้งาน")
        st.stop()
    if not tpl_file:
        st.warning("กรุณาแนบไฟล์ template.docx")
        st.stop()
    if not uploaded_file:
        st.warning("กรุณาแนบไฟล์โครงการสอน")
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
            "gemini-2.5-flash",
            "gemini-2.0-flash",
            "gemini-2.0-flash-lite"
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
        target_weekday_1 = day_map.get(day_1, 0)
        target_weekday_2 = day_map.get(day_2, 1) if is_split else None

        for idx, w in enumerate(final_weeks):
            progress_bar.progress(int(((idx + 1) / total_count) * 100))
            status_text.text(f"กำลังลงข้อมูลสัปดาห์ที่ {w.get('week')} ในแบบฟอร์มวิทยาลัย...")

            week_num = w.get("week", idx + 1)
            base_week_date = start_date + timedelta(weeks=(week_num - 1))
            date_dt_1 = base_week_date + timedelta(days=(target_weekday_1 - base_week_date.weekday()))
            date_str_1 = format_thai_date(date_dt_1)
            time_str_1 = f"เวลา {time_1}"

            if is_split:
                date_dt_2 = base_week_date + timedelta(days=(target_weekday_2 - base_week_date.weekday()))
                date_str_2 = format_thai_date(date_dt_2)
                time_str_2 = f"เวลา {time_2}"

                date_display = f"{date_str_1}\n{date_str_2}"
                time_display = f"{time_str_1}\n{time_str_2}"
            else:
                date_display = date_str_1
                time_display = time_str_1

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
                merged_doc.add_page_break()
                for el in sub_doc.element.body:
                    if el.tag.endswith('sectPr'):
                        continue
                    merged_doc.element.body.append(copy.deepcopy(el))

        output_stream = io.BytesIO()
        merged_doc.save(output_stream)
        output_stream.seek(0)

        progress_bar.progress(100)
        status_text.empty()

        st.success(f"🎉 สร้างเอกสารครบ {target_weeks} สัปดาห์ สำหรับระดับ {class_level} เรียบร้อย 100%!")
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
