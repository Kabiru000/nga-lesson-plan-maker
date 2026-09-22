import streamlit as st
import json
import io
import os
import re
import time
import fitz  # PyMuPDF
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

st.set_page_config(page_title="Noble Guide Academy Suite", layout="wide")

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

def set_cell_shading(cell, color_hex: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def clean_json_response(raw_text: str):
    text = raw_text.strip()
    if text.startswith("```json"):
        text = text[7:]
    elif text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    return json.loads(text.strip())

def extract_file_text(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
    fname = uploaded_file.name.lower()
    try:
        if fname.endswith(".pdf"):
            file_bytes = uploaded_file.read()
            uploaded_file.seek(0)
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            return "\n".join([page.get_text() for page in doc])
        elif fname.endswith(".docx"):
            d = Document(uploaded_file)
            return "\n".join([p.text for p in d.paragraphs if p.text])
        elif fname.endswith(".txt"):
            return uploaded_file.read().decode("utf-8", errors="ignore")
    except Exception as e:
        st.warning(f"Could not read {uploaded_file.name}: {e}")
    return ""

def execute_generation_with_retry(client, prompt: str):
    # Updated to active models: primary gemini-3.6-flash, fallback gemini-2.5-flash
    models_to_try = ["gemini-3.6-flash", "gemini-2.5-flash"]
    last_error = None

    for model_name in models_to_try:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        temperature=0.2,
                        response_mime_type="application/json"
                    )
                )
                return clean_json_response(response.text)
            except Exception as err:
                last_error = err
                time.sleep(2)
                continue
    raise last_error

# --- DOCX: Lesson Plans ---
def generate_docx_bytes(plans_data: list, term_label: str) -> io.BytesIO:
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

        r_plan = p_title.add_run(f"Lesson Plan - {term_label}")
        r_plan.bold = True
        r_plan.font.size = Pt(11)
        r_plan.font.name = "Times New Roman"

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
            ("Instructional/Teaching Resources:", plan.get('resources', ''), "bullet"),
            ("Reference(s):", plan.get('references', ''), "none"),
            ("Prior Knowledge and Connection:", plan.get('prior_knowledge', ''), "none"),
            ("Direct Teaching:", plan.get('direct_teaching', ''), "bullet"),
            ("Guided Practice (Students’ Active Learning):", plan.get('guided_practice', ''), "bullet"),
            ("Evaluation:", plan.get('evaluation', ''), "bullet"),
            ("Closure (Plenary):", plan.get('closure', 'Concludes the lesson by giving a neat and tidy summary of the lesson.'), "none"),
            ("Lesson Summary Notes:", plan.get('summary_notes', ''), "bullet"),
            ("Assignment:", plan.get('assignment', ''), "bullet"),
            ("HoD’s Comment And Signature:", "", "none")
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

            if heading.startswith("HoD’s Comment"):
                p_empty = cell.add_paragraph()
                p_empty.paragraph_format.space_before = Pt(18)
                p_empty.paragraph_format.space_after = Pt(18)
            else:
                lines = [line.strip() for line in str(body_text).split("\n") if line.strip()]
                for l_idx, line in enumerate(lines):
                    p_item = cell.add_paragraph()
                    p_item.paragraph_format.space_before = Pt(1)
                    p_item.paragraph_format.space_after = Pt(1)
                    p_item.paragraph_format.line_spacing = 1.15

                    clean_text = line.lstrip("0123456789.-◆*• ")

                    if bullet_type == "decimal":
                        p_item.paragraph_format.left_indent = Inches(0.2)
                        r_item = p_item.add_run(f"{l_idx + 1}. {clean_text}")
                        r_item.font.name = "Calibri"
                        r_item.font.size = Pt(10)
                    elif bullet_type == "bullet":
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

# --- DOCX: Lesson Notes ---
def format_equation_line(p, eq_str: str):
    clean_eq = eq_str.replace("->", " → ").replace("<=>", " ⇌ ")
    tokens = re.split(r'(\s+|[()+→⇌=Δ])', clean_eq)
    for token in tokens:
        if not token:
            continue
        if re.fullmatch(r'\d*[\+\-]', token):
            r = p.add_run(token)
            r.font.superscript = True
            r.font.name = "Cambria Math"
            r.font.size = Pt(11)
            r.bold = True
        elif re.search(r'[A-Z][a-z]?\d+', token):
            sub_parts = re.split(r'(\d+)', token)
            for sp in sub_parts:
                r = p.add_run(sp)
                r.font.name = "Cambria Math"
                r.font.size = Pt(11)
                if sp.isdigit():
                    r.font.subscript = True
        else:
            r = p.add_run(token)
            r.font.name = "Cambria Math"
            r.font.size = Pt(11)
            if token in ["→", "⇌", "="]:
                r.bold = True

def add_clean_paragraph(doc, text: str, bullet: bool = False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    if bullet:
        p.paragraph_format.left_indent = Inches(0.25)
        r_b = p.add_run("- ")
        r_b.font.name = "Calibri"
        r_b.font.size = Pt(10.5)
        r_b.bold = True

    parts = re.split(r'(\*\*.*?\*\*)', text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            r = p.add_run(part[2:-2])
            r.bold = True
            r.font.name = "Calibri"
            r.font.size = Pt(10.5)
        else:
            r = p.add_run(part)
            r.font.name = "Calibri"
            r.font.size = Pt(10.5)
    return p

def generate_astar_note_docx(note_data: dict) -> io.BytesIO:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    p_mast = doc.add_paragraph()
    p_mast.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_mast.paragraph_format.space_before = Pt(0)
    p_mast.paragraph_format.space_after = Pt(2)
    
    r_inst = p_mast.add_run("NOBLE GUIDE ACADEMY, ABUJA\n")
    r_inst.bold = True
    r_inst.font.name = "Arial"
    r_inst.font.size = Pt(13)
    r_inst.font.color.rgb = RGBColor(0, 51, 102)

    r_banner = p_mast.add_run("A* TARGETED COMPREHENSIVE STUDY NOTE\n")
    r_banner.bold = True
    r_banner.font.name = "Arial"
    r_banner.font.size = Pt(9.5)
    r_banner.font.color.rgb = RGBColor(110, 110, 110)

    ribbon = doc.add_table(rows=1, cols=3)
    ribbon.style = "Table Grid"
    ribbon.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_w = (Inches(2.5), Inches(3.2), Inches(1.3))
    for c, w in zip(ribbon.rows[0].cells, col_w):
        c.width = w
        set_cell_margins(c, top=60, bottom=60, left=90, right=90)
        set_cell_shading(c, "003366")

    labels = [
        f"SUBJECT: {note_data.get('subject', '').upper()}",
        f"TOPIC: {note_data.get('topic', '').upper()}",
        f"CLASS: {note_data.get('class_name', '').upper()}"
    ]
    for cell, txt in zip(ribbon.rows[0].cells, labels):
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(txt)
        r.bold = True
        r.font.name = "Arial"
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    doc.add_paragraph().paragraph_format.space_before = Pt(8)

    t_voc = doc.add_table(rows=1, cols=1)
    t_voc.style = "Table Grid"
    t_voc.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_voc = t_voc.rows[0].cells[0]
    c_voc.width = Inches(7.0)
    set_cell_margins(c_voc, top=80, bottom=80, left=120, right=120)
    set_cell_shading(c_voc, "F4F7FA")

    p_vhead = c_voc.paragraphs[0]
    p_vhead.paragraph_format.space_after = Pt(4)
    r_vh = p_vhead.add_run("🔑 EXAMINER MANDATORY MARK-SCHEME VOCABULARY")
    r_vh.bold = True
    r_vh.font.name = "Arial"
    r_vh.font.size = Pt(10)
    r_vh.font.color.rgb = RGBColor(0, 51, 102)

    for kw in note_data.get("keywords", []):
        p_k = c_voc.add_paragraph()
        p_k.paragraph_format.space_before = Pt(1)
        p_k.paragraph_format.space_after = Pt(2)
        r_term = p_k.add_run(f"• {kw.get('term', '')}: ")
        r_term.bold = True
        r_term.font.name = "Calibri"
        r_term.font.size = Pt(10)
        p_k.add_run(kw.get('meaning', '')).font.name = "Calibri"

    doc.add_paragraph().paragraph_format.space_before = Pt(6)

    p_ch = doc.add_paragraph()
    r_ch = p_ch.add_run("📘 CORE LESSON NOTES & DETAILED MECHANISMS")
    r_ch.bold = True
    r_ch.font.name = "Arial"
    r_ch.font.size = Pt(11)
    r_ch.font.color.rgb = RGBColor(0, 51, 102)

    for sec_idx, section in enumerate(note_data.get("sections", [])):
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(8)
        p_sub.paragraph_format.space_after = Pt(2)
        r_sub = p_sub.add_run(f"{sec_idx + 1}. {section.get('subheading', '')}")
        r_sub.bold = True
        r_sub.font.name = "Arial"
        r_sub.font.size = Pt(10.5)

        for pt in section.get("points", []):
            add_clean_paragraph(doc, pt, bullet=True)

        equations = section.get("equations", [])
        if equations:
            t_eq = doc.add_table(rows=1, cols=1)
            t_eq.style = "Table Grid"
            t_eq.alignment = WD_TABLE_ALIGNMENT.CENTER
            c_eq = t_eq.rows[0].cells[0]
            c_eq.width = Inches(6.8)
            set_cell_margins(c_eq, top=60, bottom=60, left=120, right=120)
            set_cell_shading(c_eq, "F9FAFB")

            p_eq_head = c_eq.paragraphs[0]
            r_eq_head = p_eq_head.add_run("Balanced Chemical & Ionic Equations:")
            r_eq_head.bold = True
            r_eq_head.font.size = Pt(9)
            r_eq_head.font.color.rgb = RGBColor(70, 70, 70)

            for eq_item in equations:
                p_eq = c_eq.add_paragraph()
                p_eq.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_eq.paragraph_format.space_before = Pt(2)
                p_eq.paragraph_format.space_after = Pt(2)
                format_equation_line(p_eq, eq_item)

    doc.add_paragraph().paragraph_format.space_before = Pt(6)
    t_warn = doc.add_table(rows=1, cols=1)
    t_warn.style = "Table Grid"
    t_warn.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_warn = t_warn.rows[0].cells[0]
    c_warn.width = Inches(7.0)
    set_cell_margins(c_warn, top=80, bottom=80, left=120, right=120)
    set_cell_shading(c_warn, "FDF3F2")

    p_whead = c_warn.paragraphs[0]
    p_whead.paragraph_format.space_after = Pt(4)
    r_wh = p_whead.add_run("⚠️ EXAMINER WARNING: COMMON MISCONCEPTIONS & LOST MARKS")
    r_wh.bold = True
    r_wh.font.name = "Arial"
    r_wh.font.size = Pt(10)
    r_wh.font.color.rgb = RGBColor(169, 50, 38)

    for pit in note_data.get("common_pitfalls", []):
        p_pit = c_warn.add_paragraph()
        p_pit.paragraph_format.space_before = Pt(2)
        p_pit.paragraph_format.space_after = Pt(2)
        p_pit.add_run("• ").bold = True
        r_p = p_pit.add_run(pit)
        r_p.font.name = "Calibri"
        r_p.font.size = Pt(10)

    if note_data.get("worked_example"):
        doc.add_paragraph().paragraph_format.space_before = Pt(6)
        t_work = doc.add_table(rows=1, cols=1)
        t_work.style = "Table Grid"
        t_work.alignment = WD_TABLE_ALIGNMENT.CENTER
        c_work = t_work.rows[0].cells[0]
        c_work.width = Inches(7.0)
        set_cell_margins(c_work, top=80, bottom=80, left=120, right=120)
        set_cell_shading(c_work, "F9FBFD")

        p_workhead = c_work.paragraphs[0]
        p_workhead.paragraph_format.space_after = Pt(4)
        r_wkh = p_workhead.add_run("📝 A* STEP-BY-STEP WORKED EXAMPLE")
        r_wkh.bold = True
        r_wkh.font.name = "Arial"
        r_wkh.font.size = Pt(10)
        r_wkh.font.color.rgb = RGBColor(0, 51, 102)

        p_wq = c_work.add_paragraph()
        r_wq_title = p_wq.add_run("Question: ")
        r_wq_title.bold = True
        p_wq.add_run(note_data['worked_example'].get('question', '')).font.name = "Calibri"

        for s_idx, stp in enumerate(note_data['worked_example'].get('solution_steps', [])):
            p_stp = c_work.add_paragraph()
            p_stp.paragraph_format.left_indent = Inches(0.2)
            r_st_title = p_stp.add_run(f"Step {s_idx + 1}: ")
            r_st_title.bold = True
            format_equation_line(p_stp, stp)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# --- AUTH & SIDEBAR ---
query_params = st.query_params
saved_key = query_params.get("k", "")

with st.sidebar:
    st.header("🔑 API Configuration")
    user_key = st.text_input("Gemini API Key:", value=saved_key, type="password", help="Free key from aistudio.google.com")
    if st.button("💾 Save Key to this Device"):
        if user_key.strip():
            st.query_params["k"] = user_key.strip()
            st.success("Key saved!")
        else:
            st.query_params.clear()
            st.info("Key cleared.")
    st.markdown("---")
    st.markdown("**Quota Reset (Free):**\n1. Go to [aistudio.google.com/apikey](https://aistudio.google.com/apikey)\n2. Click **Create API key in new project**\n3. Paste the key above and Save.")

api_key = user_key.strip()
if not api_key:
    st.warning("👈 Enter your Gemini API key in the left sidebar to begin.")
    st.stop()

if "plans_output" not in st.session_state:
    st.session_state.plans_output = None
if "plans_filename" not in st.session_state:
    st.session_state.plans_filename = ""
if "notes_output" not in st.session_state:
    st.session_state.notes_output = None
if "notes_filename" not in st.session_state:
    st.session_state.notes_filename = ""

tab1, tab2 = st.tabs(["📋 Inspectorate Lesson Plans", "🌟 A* Lesson Notes"])

# =======================================================
# TAB 1: LESSON PLAN GENERATOR
# =======================================================
with tab1:
    st.subheader("🗓️ Session & Term Scheduling")
    s_col1, s_col2, s_col3 = st.columns(3)
    with s_col1:
        term_selected = st.selectbox("Select Term", ["Term 1 (First Term)", "Term 2 (Second Term)", "Term 3 (Third Term)"], key="lp_term")
    with s_col2:
        week_selected = st.selectbox("Select Week", [f"Week {i}" for i in range(1, 9)], key="lp_week")
    with s_col3:
        lesson_count = st.number_input("Contacts/Lessons this Week", min_value=1, max_value=5, value=3, step=1, key="lp_count")

    st.subheader("📋 Teacher & Class Information")
    c1, c2 = st.columns(2)
    with c1:
        staff_name = st.text_input("Teacher's Full Name", value="AMINU KABIRU", key="lp_staff")
        subject = st.text_input("Subject", value="CHEMISTRY", key="lp_subj")
        class_name = st.text_input("Class", value="Year 11", key="lp_cls")
    with c2:
        date_schedule = st.text_input("Date(s) for this Week", value="7th & 8th June, 2026", key="lp_dates")
        duration = st.text_input("Duration of Lesson", value="50minutes", key="lp_dur")
        no_in_class = st.text_input("Number in Class", value="20", key="lp_num")

    c3, c4 = st.columns(2)
    with c3:
        unit_topic = st.text_input("Unit Topic for this Week", value="Acids, Bases and Salts", key="lp_utop")
    with c4:
        textbooks = st.text_input("Reference Textbooks", value="New School Chemistry & Cambridge IGCSE Chemistry", key="lp_tb")

    st.subheader(f"📖 Specific Lesson Topics for {lesson_count} Contact(s)")
    contact_topics = {}
    topic_cols = st.columns(int(lesson_count))
    for i in range(int(lesson_count)):
        with topic_cols[i]:
            contact_topics[f"Lesson {i+1}"] = st.text_input(f"Lesson {i+1} Topic:", placeholder=f"Sub-topic {i+1}", key=f"lp_top_{i+1}")

    st.subheader("📎 Curriculum, Scheme & Note Uploads (Optional)")
    u_col1, u_col2, u_col3 = st.columns(3)
    with u_col1:
        curriculum_file = st.file_uploader("1. Curriculum Framework [.pdf, .docx, .txt]", type=["pdf", "docx", "txt"], key="lp_cur")
    with u_col2:
        scheme_file = st.file_uploader("2. Annual Scheme of Work [.pdf, .docx, .txt]", type=["pdf", "docx", "txt"], key="lp_sch")
    with u_col3:
        notes_file = st.file_uploader("3. Lesson Notes / Textbook [.pdf, .docx, .txt]", type=["pdf", "docx", "txt"], key="lp_not")

    scheme_detail = st.text_area(
        f"Curriculum Objectives & Codes for {term_selected}, {week_selected} (Paste here if not uploading files):",
        height=90,
        placeholder="CHE1.1.1 Identify Cations and Anions in a solution.\nCHE1.1.2 Test for aqueous cations using sodium hydroxide.",
        key="lp_manual_objs"
    )

    if st.button("Generate Inspection Plans", type="primary", key="btn_gen_plans"):
        extracted_curriculum = extract_file_text(curriculum_file)
        extracted_scheme = extract_file_text(scheme_file)
        extracted_notes = extract_file_text(notes_file)

        if not scheme_detail.strip() and not extracted_scheme and not extracted_curriculum:
            st.warning(f"Please provide objectives or upload your Scheme/Curriculum for {term_selected} {week_selected}.")
        else:
            with st.spinner(f"Generating {lesson_count} inspection lesson plan(s)..."):
                combined_docs = scheme_detail.strip()
                if extracted_curriculum:
                    combined_docs += f"\n\n--- CURRICULUM FRAMEWORK ---\n{extracted_curriculum[:8000]}"
                if extracted_scheme:
                    combined_docs += f"\n\n--- SCHEME OF WORK ({term_selected.upper()} {week_selected.upper()}) ---\n{extracted_scheme[:12000]}"
                if extracted_notes:
                    combined_docs += f"\n\n--- REFERENCE NOTES ---\n{extracted_notes[:5000]}"

                topic_instructions = []
                for k, v in contact_topics.items():
                    val = v.strip()
                    if val:
                        topic_instructions.append(f"{k}: Use exact topic '{val}'")
                    else:
                        topic_instructions.append(f"{k}: Derive appropriate specific sub-topic from the scheme")
                topic_summary = "\n".join(topic_instructions)

                prompt = f"""
                You are an expert Inspectorate Curriculum Specialist for Noble Guide Academy, Abuja.
                Generate exactly {lesson_count} sequential lesson plans (Lesson 1 to Lesson {lesson_count}) as a valid JSON array.

                TARGET TIME FRAME:
                - Term: {term_selected}
                - Week: {week_selected}
                - Number of Lessons/Contacts: {lesson_count}

                METADATA:
                - Staff Name: {staff_name}
                - Subject: {subject}
                - Class: {class_name}
                - Week: {week_selected}
                - Date: {date_schedule}
                - Duration: {duration}
                - No. in Class: {no_in_class}
                - Unit Topic: {unit_topic}
                - Reference Books: {textbooks}

                EXACT LESSON TOPIC ASSIGNMENTS:
                {topic_summary}

                CURRICULUM SPECIFICATIONS, SCHEME OF WORK & REFERENCE MATERIAL:
                {combined_docs}

                CRITICAL PEDAGOGICAL & FORMATTING REQUIREMENTS:
                1. STRICT CURRICULUM FRAMEWORK FIDELITY:
                   - 'direct_teaching': Must be derived strictly and exclusively from the uploaded Curriculum Framework and Scheme of Work.
                   - 'prior_knowledge': Must be derived strictly from preceding stages in the curriculum framework, formatted as "The students are familiar with...".
                2. LESSON OBJECTIVES: Active Bloom's verbs citing matching framework code (e.g. [CHE1.1.1]).
                3. GUIDED PRACTICE: EVERY individual active learning step MUST begin with its matching curriculum framework code, e.g., "[CHE1.1.1] Students test unknown solutions...".
                4. INSTRUCTIONAL RESOURCES & WORKING URL LINKS:
                   - Direct clickable YouTube search URL: "YouTube Video: [Title] - [https://www.youtube.com/results?search_query=](https://www.youtube.com/results?search_query=)[encoded+terms]"
                   - Direct clickable PhET simulation URL: "PhET Simulation: [Title] - [https://phet.colorado.edu/en/simulations/filter?subjects=](https://phet.colorado.edu/en/simulations/filter?subjects=)[topic]&type=html"
                5. LESSON SUMMARY NOTES: 3-4 bullet points summarizing key concepts (will appear directly below Closure).
                6. CLOSURE (PLENARY): Exactly "Concludes the lesson by giving a neat and tidy summary of the lesson.".
                7. HOD SECTION: The "hod_comment" field must be empty ("").
                8. Return strictly a JSON array containing exactly {lesson_count} objects.

                JSON Schema keys per item:
                "school_name", "staff_name", "subject", "unit_topic", "lesson_topic", "date", "period", "week", "lesson_number", "sex", "duration", "class_name", "no_in_class", "objectives", "resources", "references", "prior_knowledge", "direct_teaching", "guided_practice", "evaluation", "closure", "summary_notes", "assignment", "hod_comment"
                """
                try:
                    client = genai.Client(api_key=api_key)
                    data = execute_generation_with_retry(client, prompt)
                    docx_bytes = generate_docx_bytes(data, term_selected)
                    clean_term = term_selected.replace(' ', '_')
                    clean_week = week_selected.replace(' ', '_')
                    st.session_state.plans_output = docx_bytes
                    st.session_state.plans_filename = f"NGA_{clean_term}_{clean_week}_{subject}.docx"
                except Exception as e:
                    st.error(f"Generation error: {e}")

    if st.session_state.plans_output:
        st.success("Lesson plan document ready!")
        st.download_button(
            label="📥 Download Inspection Word Document (.docx)",
            data=st.session_state.plans_output,
            file_name=st.session_state.plans_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_btn_plans"
        )

# =======================================================
# TAB 2: A* LESSON NOTE GENERATOR
# =======================================================
with tab2:
    st.subheader("🌟 Generate A* Cambridge/WAEC-Grade Lesson Notes")
    st.caption("Upload textbook extracts or notes to generate examiner-targeted study notes with equations, keywords, and worked examples.")

    an_c1, an_c2 = st.columns(2)
    with an_c1:
        an_subject = st.text_input("Subject", value="CHEMISTRY", key="note_subj")
        an_topic = st.text_input("Specific Topic", value="Acids, Bases and Salts", key="note_topic")
        an_class = st.text_input("Class Target", value="Year 11 / Cambridge IGCSE", key="note_class")
    with an_c2:
        textbook_file = st.file_uploader("Upload Textbook / Chapter PDF [.pdf, .docx, .txt]", type=["pdf", "docx", "txt"], key="note_file")

    an_extra_details = st.text_area(
        "Target Curriculum Objectives / Codes (Optional):",
        placeholder="e.g. Focus on preparation of insoluble salts via precipitation, state symbols, and ionic equations.",
        height=80,
        key="note_extra"
    )

    if st.button("Generate A* Comprehensive Lesson Note", type="primary", key="btn_gen_notes"):
        extracted_textbook = extract_file_text(textbook_file)
        with st.spinner(f"Compiling A* comprehensive revision note for '{an_topic}'..."):
            note_prompt = f"""
            You are a Senior Principal Examiner preparing a publication-grade, A* student revision note.
            SUBJECT: {an_subject}
            TOPIC: {an_topic}
            CLASS: {an_class}
            CURRICULUM SPECIFICATIONS: {an_extra_details}

            PRIMARY SOURCE EXTRACT:
            \"\"\"
            {extracted_textbook[:10000] if extracted_textbook else "Strictly align with Cambridge IGCSE / WAEC curriculum specifications."}
            \"\"\"

            CRITICAL FIDELITY REQUIREMENTS:
            1. STRICT GROUNDING: Derive definitions, chemical explanations, and procedures directly from the source text.
            2. PROFESSIONAL EQUATION FORMATTING:
               - Place all balanced equations in the dedicated "equations" array under each section.
               - Write complete equations with correct state symbols (s, l, g, aq).
               - Example: "CaCO3(s) + 2HCl(aq) -> CaCl2(aq) + H2O(l) + CO2(g)"
            3. CLEAN STRUCTURE: In the "points" array, write concise, informative sentences. Do not use raw markdown asterisks (**).
            4. EXAMINER MANDATORY VOCABULARY: 4 to 6 exact technical keywords defined in the text.
            5. COMMON PITFALLS: 3 specific errors where students frequently lose marks.
            6. WORKED EXAMPLE: A multi-step calculation or preparation problem with numbered steps.

            JSON Schema:
            {{
                "subject": "{an_subject}",
                "topic": "{an_topic}",
                "class_name": "{an_class}",
                "keywords": [{{"term": "string", "meaning": "string"}}],
                "sections": [
                    {{
                        "subheading": "string",
                        "points": ["string", "string"],
                        "equations": ["string"]
                    }}
                ],
                "common_pitfalls": ["string", "string"],
                "worked_example": {{
                    "question": "string",
                    "solution_steps": ["string", "string"]
                }}
            }}
            """
            try:
                client = genai.Client(api_key=api_key)
                note_json = execute_generation_with_retry(client, note_prompt)
                note_docx = generate_astar_note_docx(note_json)
                st.session_state.notes_output = note_docx
                st.session_state.notes_filename = f"AStar_Note_{an_topic.replace(' ', '_')}.docx"
            except Exception as e:
                st.error(f"Generation error: {e}")

    if st.session_state.notes_output:
        st.success("A* Lesson Note compiled successfully!")
        st.download_button(
            label="📥 Download A* Word Document (.docx)",
            data=st.session_state.notes_output,
            file_name=st.session_state.notes_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_btn_notes"
        )
