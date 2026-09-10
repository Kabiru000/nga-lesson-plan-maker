import streamlit as st
import json
import urllib.request
import ssl
import io
import os
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

st.set_page_config(page_title="Noble Guide Academy Lesson Plan Generator", layout="centered")

def set_cell_margins(cell, top=70, bottom=70, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def generate_docx_bytes(plans_data: list) -> io.BytesIO:
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    logo_path = os.path.join(os.path.dirname(__file__), "NGA_logo.png")
    if not os.path.exists(logo_path):
        logo_path = "NGA_logo.png"

    for idx, plan in enumerate(plans_data):
        if idx > 0:
            doc.add_page_break()

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(0)
        p_title.paragraph_format.space_after = Pt(2)

        if os.path.exists(logo_path):
            r_logo = p_title.add_run()
            r_logo.add_picture(logo_path, width=Inches(0.57), height=Inches(0.52))
            p_title.add_run("\n")

        r_school = p_title.add_run("Noble Guide Academy, Abuja\n")
        r_school.bold = True
        r_school.font.size = Pt(13)
        r_school.font.name = "Times New Roman"

        r_plan = p_title.add_run("Lesson Plan")
        r_plan.bold = True
        r_plan.font.size = Pt(11)
        r_plan.font.name = "Times New Roman"

        # 1-Row, 3-Column Header Table
        meta_table = doc.add_table(rows=1, cols=3)
        meta_table.style = "Table Grid"
        meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        meta_table.autofit = False

        col_widths = (Inches(2.8), Inches(2.2), Inches(2.2))
        for cell, w in zip(meta_table.rows[0].cells, col_widths):
            cell.width = w

        col1_text = (
            f"Staff’s Name: {plan.get('staff_name', '')}\n"
            f"Subject:  {plan.get('subject', '')}\n"
            f"Unit Topic: {plan.get('unit_topic', '')}\n"
            f"Lesson Topic: {plan.get('lesson_topic', '')}"
        )
        col2_text = (
            f"Date: {plan.get('date', '')}\n"
            f"Period: {plan.get('period', '')}\n"
            f"Week: {plan.get('week', '')}\n"
            f"Lesson: {plan.get('lesson_number', '')}"
        )
        col3_text = (
            f"Sex: {plan.get('sex', 'Mixed')}\n"
            f"Duration: {plan.get('duration', '50minutes')}\n"
            f"Class: {plan.get('class_name', '')}\n"
            f"No. in class: {plan.get('no_in_class', '20')}"
        )

        for cell, text in zip(meta_table.rows[0].cells, [col1_text, col2_text, col3_text]):
            cell.text = text
            set_cell_margins(cell, top=60, bottom=60, left=90, right=90)
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.15
                for r in p.runs:
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(9.5)

        p_spacer = doc.add_paragraph()
        p_spacer.paragraph_format.space_before = Pt(3)
        p_spacer.paragraph_format.space_after = Pt(3)

        sections = [
            ("Lesson Objectives: By the end of the lesson students, should be able to:", plan.get('objectives', ''), "decimal"),
            ("Instructional/Teaching Resources:", plan.get('resources', ''), "hyphen"),
            ("Reference(s):", plan.get('references', ''), "none"),
            ("Prior Knowledge and Connection:", plan.get('prior_knowledge', ''), "none"),
            ("Direct Teaching:", plan.get('direct_teaching', ''), "diamond"),
            ("Guided Practice (Students’ Active Learning):", plan.get('guided_practice', ''), "diamond"),
            ("Evaluation:", plan.get('evaluation', ''), "hyphen"),
            ("Closure (Plenary):", plan.get('closure', 'Concludes the lesson by giving a neat and tidy summary of the lesson.'), "none"),
            ("Assignment:", plan.get('assignment', ''), "hyphen"),
            ("HoD’s Comment And Signature:", plan.get('hod_comment', ''), "none")
        ]

        for heading, body_text, bullet_type in sections:
            box_table = doc.add_table(rows=1, cols=1)
            box_table.style = "Table Grid"
            box_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            box_table.autofit = False

            cell = box_table.rows[0].cells[0]
            cell.width = Inches(7.2)
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)

            p_head = cell.paragraphs[0]
            p_head.paragraph_format.space_before = Pt(0)
            p_head.paragraph_format.space_after = Pt(2)
            r_head = p_head.add_run(heading)
            r_head.bold = True
            r_head.font.size = Pt(10)
            r_head.font.name = "Times New Roman"

            lines = [line.strip() for line in str(body_text).split("\n") if line.strip()]
            for l_idx, line in enumerate(lines):
                p_item = cell.add_paragraph()
                p_item.paragraph_format.space_before = Pt(1)
                p_item.paragraph_format.space_after = Pt(1)
                p_item.paragraph_format.line_spacing = 1.15

                clean_text = line.lstrip("0123456789.-◆* ")

                if bullet_type == "decimal":
                    p_item.paragraph_format.left_indent = Inches(0.2)
                    r_item = p_item.add_run(f"{l_idx + 1}. {clean_text}")
                    r_item.font.name = "Calibri"
                    r_item.font.size = Pt(10)
                elif bullet_type == "diamond":
                    p_item.paragraph_format.left_indent = Inches(0.2)
                    r_item = p_item.add_run(f"◆ {clean_text}")
                    r_item.font.name = "Calibri"
                    r_item.font.size = Pt(10)
                elif bullet_type == "hyphen":
                    p_item.paragraph_format.left_indent = Inches(0.2)
                    r_item = p_item.add_run(f"- {clean_text}")
                    r_item.font.name = "Calibri"
                    r_item.font.size = Pt(10)
                else:
                    r_item = p_item.add_run(clean_text)
                    r_item.font.name = "Times New Roman"
                    r_item.font.size = Pt(9.5)

            p_gap = doc.add_paragraph()
            p_gap.paragraph_format.space_before = Pt(2)
            p_gap.paragraph_format.space_after = Pt(2)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

def call_gemini(prompt: str, key: str) -> str:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={key}"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.2,
            "maxOutputTokens": 4096,
            "responseMimeType": "application/json"
        }
    }
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req, context=ctx, timeout=60) as resp:
        res_data = json.loads(resp.read().decode("utf-8"))
        parts = res_data["candidates"][0]["content"]["parts"]
        return "".join([p["text"] for p in parts if "text" in p]).strip()

st.title("📚 Noble Guide Academy Lesson Plan Generator")
st.markdown("Fast, inspectorate-grade lesson plan generation for Noble Guide Academy.")

try:
    api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    api_key = ""

if not api_key:
    st.error("API key is not configured in Streamlit Secrets.")
    st.stop()

col1, col2 = st.columns(2)
with col1:
    staff_name = st.text_input("Teacher's Full Name", value="AMINU KABIRU")
    subject = st.text_input("Subject", value="CHEMISTRY")
    class_name = st.text_input("Class", value="Year 11")
with col2:
    week = st.text_input("Week Number", value="6")
    unit_topic = st.text_input("Unit Topic", value="Chemical and investigations")
    textbooks = st.text_input("Reference Textbooks", value="New School Chemistry and Cambridge Chemistry Syllabus IGCSE Course Book")

scheme_detail = st.text_area(
    "Scheme Objectives & Curriculum Codes (One per line)",
    height=140,
    placeholder="CHE1.1.1 Identify Cations and Anions in a solution.\nCHE1.1.2 Test for aqueous cations using sodium hydroxide."
)

if st.button("Generate Inspection Plans", type="primary"):
    if not scheme_detail.strip():
        st.warning("Please enter at least one curriculum objective.")
    else:
        with st.spinner("Generating 4 lesson plans (typically takes ~10 seconds)..."):
            prompt = f"""
            Generate exactly 4 structured lesson plans (Lesson 1, 2, 3, and 4) as a valid JSON array for Noble Guide Academy:
            Teacher: {staff_name} | Subject: {subject} | Class: {class_name} | Week: {week} | Unit Topic: {unit_topic} | Books: {textbooks}
            Curriculum Objectives:
            {scheme_detail}

            Rules:
            1. Output must be a pure JSON array with exactly 4 objects.
            2. 'resources': Include 1 specific YouTube search recommendation (e.g., FuseSchool, Cognito) and 1 simulation/lab aid (e.g., PhET Interactive Simulations). Each on a new line.
            3. 'objectives': 2-3 concise Bloom's action-verb objectives, each on a new line.
            4. 'direct_teaching': 2-3 direct instruction steps, each on a new line.
            5. 'guided_practice': 2-3 active learning steps, each on a new line.
            6. 'evaluation': 1-2 assessment questions, each on a new line.
            7. 'assignment': 1-2 homework questions, each on a new line.
            8. 'prior_knowledge': format "The students are familiar with...".
            9. 'closure': "Concludes the lesson by giving a neat and tidy summary of the lesson.".

            JSON fields required per item:
            "school_name", "staff_name", "subject", "unit_topic", "lesson_topic", "date", "period", "week", "lesson_number", "sex", "duration", "class_name", "no_in_class", "objectives", "resources", "references", "prior_knowledge", "direct_teaching", "guided_practice", "evaluation", "closure", "assignment", "hod_comment"
            """
            try:
                raw_json = call_gemini(prompt, api_key)
                data = json.loads(raw_json)
                file_data = generate_docx_bytes(data)

                st.success("Lesson plans ready!")
                st.download_button(
                    label="📥 Download Word Document (.docx)",
                    data=file_data,
                    file_name=f"NGA_Lesson_Plan_Week_{week}_{unit_topic.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except urllib.error.URLError as e:
                st.error(f"Network Timeout: {e}")
            except Exception as e:
                st.error(f"Generation error: {e}")
