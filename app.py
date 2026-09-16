import streamlit as st
import json
import io
import os
import re
import time
from pypdf import PdfReader
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

st.set_page_config(page_title="Noble Guide Academy Suite", layout="wide")

def clean_json_response(raw_text: str):
    """Strip markdown code block wrappers if present."""
    text = raw_text.strip()
    if text.startswith("```json"):
        text = text[7:]
    elif text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    return json.loads(text.strip())

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

def extract_file_text(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
    fname = uploaded_file.name.lower()
    try:
        if fname.endswith(".pdf"):
            reader = PdfReader(uploaded_file)
            return "\n".join([page.extract_text() or "" for page in reader.pages])
        elif fname.endswith(".docx"):
            d = Document(uploaded_file)
            return "\n".join([p.text for p in d.paragraphs if p.text])
        elif fname.endswith(".txt"):
            return uploaded_file.read().decode("utf-8", errors="ignore")
    except Exception as e:
        st.warning(f"Could not read {uploaded_file.name}: {e}")
    return ""

def execute_generation_with_retry(client, prompt: str):
    last_error = None
    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model="gemini-3.6-flash",
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=0.2,
                    response_mime_type="application/json"
                )
            )
            return clean_json_response(response.text)
        except Exception as err:
            err_str = str(err)
            last_error = err
            if any(k in err_str for k in ["503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED"]):
                time.sleep(3 * (attempt + 1))
                continue
            raise err
    raise last_error

# --- Matplotlib Visual Graph Generator ---
def generate_sample_topic_graph(graph_type: str) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=(6, 3.5), dpi=150)
    fig.patch.set_facecolor('#ffffff')
    ax.set_facecolor('#fafafa')

    if "rate" in graph_type.lower():
        t = np.linspace(0, 100, 200)
        c1 = 100 * (1 - np.exp(-0.05 * t))
        c2 = 100 * (1 - np.exp(-0.02 * t))
        ax.plot(t, c1, 'b-', label='High Concentration / With Catalyst', linewidth=2)
        ax.plot(t, c2, 'r--', label='Lower Concentration / No Catalyst', linewidth=2)
        ax.set_title("Rate of Reaction: Volume of Gas vs Time", fontsize=11, fontweight='bold')
        ax.set_xlabel("Time (seconds)", fontsize=9)
        ax.set_ylabel("Volume of Gas (cm³)", fontsize=9)
        ax.legend(fontsize=8)
        ax.grid(True, linestyle=':', alpha=0.6)
    elif "heating" in graph_type.lower() or "cooling" in graph_type.lower():
        time_pts = [0, 2, 5, 8, 11, 14]
        temps = [20, 80, 80, 120, 120, 140]
        ax.plot(time_pts, temps, 'darkorange', linewidth=2.2, marker='o')
        ax.set_title("Heating Curve (Phase Changes at Constant Temperature)", fontsize=11, fontweight='bold')
        ax.set_xlabel("Time of Heating (min)", fontsize=9)
        ax.set_ylabel("Temperature (°C)", fontsize=9)
        ax.annotate('Melting Point', xy=(3.5, 80), xytext=(2, 95),
                    arrowprops=dict(facecolor='black', shrink=0.05, width=1, headwidth=5))
        ax.annotate('Boiling Point', xy=(9.5, 120), xytext=(8, 133),
                    arrowprops=dict(facecolor='black', shrink=0.05, width=1, headwidth=5))
        ax.grid(True, linestyle=':', alpha=0.6)
    else:
        x = np.linspace(0, 10, 100)
        y = np.sin(x)
        ax.plot(x, y, 'purple', linewidth=2)
        ax.set_title("Scientific Trend Analysis", fontsize=11, fontweight='bold')
        ax.set_xlabel("Independent Variable", fontsize=9)
        ax.set_ylabel("Dependent Variable", fontsize=9)
        ax.grid(True, linestyle=':', alpha=0.6)

    plt.tight_layout()
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png')
    plt.close(fig)
    img_buf.seek(0)
    return img_buf

# --- DOCX Builder: Lesson Plans ---
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

# --- DOCX Builder: A* Comprehensive Lesson Notes ---
def generate_astar_note_docx(note_data: dict, graph_type: str = "none") -> io.BytesIO:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_school = p_head.add_run("NOBLE GUIDE ACADEMY, ABUJA\n")
    r_school.bold = True
    r_school.font.size = Pt(14)
    r_school.font.name = "Times New Roman"

    r_tag = p_head.add_run("A* TARGETED COMPREHENSIVE REVISION NOTE\n")
    r_tag.bold = True
    r_tag.font.size = Pt(11)
    r_tag.font.color.rgb = RGBColor(0, 51, 102)

    r_topic = p_head.add_run(f"Subject: {note_data.get('subject','')} | Topic: {note_data.get('topic','')}\nClass: {note_data.get('class_name','')}")
    r_topic.font.size = Pt(10)
    r_topic.font.italic = True

    # 1. Vocabulary Table
    t1 = doc.add_table(rows=1, cols=1)
    t1.style = "Table Grid"
    c1 = t1.rows[0].cells[0]
    set_cell_margins(c1, top=80, bottom=80, left=120, right=120)
    p_khead = c1.paragraphs[0]
    r_khead = p_khead.add_run("🔑 EXAMINER MANDATORY MARK-SCHEME VOCABULARY")
    r_khead.bold = True
    r_khead.font.size = Pt(10.5)
    r_khead.font.color.rgb = RGBColor(0, 51, 102)

    for kw in note_data.get("keywords", []):
        p_k = c1.add_paragraph()
        p_k.paragraph_format.left_indent = Inches(0.2)
        r_term = p_k.add_run(f"• {kw.get('term', '')}: ")
        r_term.bold = True
        p_k.add_run(kw.get('meaning', ''))

    doc.add_paragraph().paragraph_format.space_before = Pt(4)

    # 2. Core Concepts
    p_c = doc.add_paragraph()
    r_chead = p_c.add_run("📘 COMPREHENSIVE CONCEPT BREAKDOWN & MECHANISMS")
    r_chead.bold = True
    r_chead.font.size = Pt(11)

    for section in note_data.get("sections", []):
        p_shead = doc.add_paragraph()
        p_shead.paragraph_format.space_before = Pt(4)
        r_sh = p_shead.add_run(f"■ {section.get('subheading', '')}")
        r_sh.bold = True
        r_sh.font.size = Pt(10.5)

        for bp in section.get("points", []):
            p_bp = doc.add_paragraph()
            p_bp.paragraph_format.left_indent = Inches(0.2)
            r_bp = p_bp.add_run(f"- {bp}")
            r_bp.font.size = Pt(10)
            r_bp.font.name = "Calibri"

    # Optional Graph Embed
    if graph_type != "none":
        doc.add_paragraph().paragraph_format.space_before = Pt(4)
        p_ghead = doc.add_paragraph()
        r_gh = p_ghead.add_run("📈 ESSENTIAL GRAPH PROFILE & INTERPRETATION")
        r_gh.bold = True
        r_gh.font.size = Pt(11)

        graph_buf = generate_sample_topic_graph(graph_type)
        doc.add_picture(graph_buf, width=Inches(5.8))
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_cap = p_cap.add_run(f"Figure: Standard Graphical Analysis for {note_data.get('topic','')}")
        r_cap.font.size = Pt(8.5)
        r_cap.font.italic = True

    doc.add_paragraph().paragraph_format.space_before = Pt(4)

    # 3. Pitfalls Table
    t_pit = doc.add_table(rows=1, cols=1)
    t_pit.style = "Table Grid"
    c_pit = t_pit.rows[0].cells[0]
    set_cell_margins(c_pit, top=80, bottom=80, left=120, right=120)
    p_phead = c_pit.paragraphs[0]
    r_ph = p_phead.add_run("⚠️ EXAMINER WARNING: COMMON MISCONCEPTIONS & LOST MARKS")
    r_ph.bold = True
    r_ph.font.size = Pt(10.5)
    r_ph.font.color.rgb = RGBColor(180, 0, 0)

    for pit in note_data.get("common_pitfalls", []):
        p_item = c_pit.add_paragraph()
        p_item.paragraph_format.left_indent = Inches(0.2)
        r_it = p_item.add_run(f"• {pit}")
        r_it.font.size = Pt(9.5)
        r_it.font.name = "Calibri"

    doc.add_paragraph().paragraph_format.space_before = Pt(4)

    # 4. Worked Example
    if note_data.get("worked_example"):
        t_w = doc.add_table(rows=1, cols=1)
        t_w.style = "Table Grid"
        c_w = t_w.rows[0].cells[0]
        set_cell_margins(c_w, top=80, bottom=80, left=120, right=120)
        p_whead = c_w.paragraphs[0]
        r_wh = p_whead.add_run("📝 A* MODEL WORKED EXAMPLE & STEP-BY-STEP SOLUTION")
        r_wh.bold = True
        r_wh.font.size = Pt(10.5)

        p_q = c_w.add_paragraph()
        r_qt = p_q.add_run(f"Question: {note_data['worked_example'].get('question','')}")
        r_qt.font.size = Pt(9.5)
        r_qt.bold = True

        for step in note_data['worked_example'].get('solution_steps', []):
            p_st = c_w.add_paragraph()
            p_st.paragraph_format.left_indent = Inches(0.2)
            p_st.add_run(f"Step: {step}").font.size = Pt(9.5)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# --- SIDEBAR API MANAGEMENT ---
query_params = st.query_params
saved_key = query_params.get("k", "")

with st.sidebar:
    st.header("🔑 API Configuration")
    user_key = st.text_input("Gemini API Key:", value=saved_key, type="password", help="Paste your free API key from aistudio.google.com")
    if st.button("💾 Save Key to this Device"):
        if user_key.strip():
            st.query_params["k"] = user_key.strip()
            st.success("Key saved! Bookmark this URL.")
        else:
            st.query_params.clear()
            st.info("Key cleared.")
    st.markdown("---")
    st.markdown("**Get a Free Key (100% Free):**\n1. Go to [aistudio.google.com](https://aistudio.google.com)\n2. Click **Get API key**\n3. Paste here and Save.")

api_key = user_key.strip()

if not api_key:
    st.warning("👈 Please enter your free Gemini API key in the left sidebar to access the suite.")
    st.stop()

# --- INITIALIZE SESSION STATE ---
if "note_docx" not in st.session_state:
    st.session_state.note_docx = None
if "note_filename" not in st.session_state:
    st.session_state.note_filename = ""
if "plans_docx" not in st.session_state:
    st.session_state.plans_docx = None
if "plans_filename" not in st.session_state:
    st.session_state.plans_filename = ""

# --- TAB NAVIGATION ---
tab_plans, tab_notes = st.tabs(["📋 Lesson Plan Generator", "🌟 A* Comprehensive Lesson Notes"])

# =======================================================
# TAB 1: LESSON PLAN GENERATOR
# =======================================================
with tab_plans:
    st.subheader("🗓️ Session & Term Scheduling")
    s_col1, s_col2, s_col3 = st.columns(3)
    with s_col1:
        term_selected = st.selectbox("Select Term", ["Term 1 (First Term)", "Term 2 (Second Term)", "Term 3 (Third Term)"], key="lp_term")
    with s_col2:
        week_selected = st.selectbox("Select Week", [f"Week {i}" for i in range(1, 9)], key="lp_week")
    with s_col3:
        lesson_count = st.number_input("Contacts/Lessons for this Subject this Week", min_value=1, max_value=5, value=3, step=1, key="lp_contacts")

    st.subheader("📋 Teacher & Class Information")
    c1, c2 = st.columns(2)
    with c1:
        staff_name = st.text_input("Teacher's Full Name", value="AMINU KABIRU", key="lp_staff")
        subject = st.text_input("Subject", value="CHEMISTRY", key="lp_sub")
        class_name = st.text_input("Class", value="Year 11", key="lp_class")
    with c2:
        date_schedule = st.text_input("Date(s) for this Week", value="7th & 8th June, 2026", key="lp_dates")
        duration = st.text_input("Duration of Lesson", value="50minutes", key="lp_dur")
        no_in_class = st.text_input("Number in Class", value="20", key="lp_num")

    c3, c4 = st.columns(2)
    with c3:
        unit_topic = st.text_input("Unit Topic for this Week", value="Chemical and investigations", key="lp_utopic")
    with c4:
        textbooks = st.text_input("Reference Textbooks", value="New School Chemistry and Cambridge Chemistry Syllabus IGCSE Course Book", key="lp_books")

    st.subheader(f"📖 Specific Lesson Topics for {lesson_count} Contact(s)")
    contact_topics = {}
    topic_cols = st.columns(int(lesson_count))
    for i in range(int(lesson_count)):
        with topic_cols[i]:
            contact_topics[f"Lesson {i+1}"] = st.text_input(
                f"Lesson {i+1} Topic:",
                placeholder=f"e.g., Sub-topic {i+1}",
                key=f"topic_contact_{i+1}"
            )

    st.subheader("📎 Curriculum, Scheme & Note Uploads (Optional)")
    u_col1, u_col2, u_col3 = st.columns(3)
    with u_col1:
        curriculum_file = st.file_uploader("1. Curriculum Framework [.pdf, .docx, .txt]", type=["pdf", "docx", "txt"], key="lp_cur")
    with u_col2:
        scheme_file = st.file_uploader("2. Annual Scheme of Work [.pdf, .docx, .txt]", type=["pdf", "docx", "txt"], key="lp_sch")
    with u_col3:
        notes_file = st.file_uploader("3. Lesson Notes / Textbook [.pdf, .docx, .txt]", type=["pdf", "docx", "txt"], key="lp_not")

    scheme_detail = st.text_area(
        f"Curriculum Objectives & Codes for {term_selected}, {week_selected} (Optional: paste here if not uploading files):",
        height=100,
        placeholder="CHE1.1.1 Identify Cations and Anions in a solution.\nCHE1.1.2 Test for aqueous cations using sodium hydroxide.",
        key="lp_manual_objs"
    )

    if st.button("Generate Inspection Plans", type="primary", key="btn_gen_plans"):
        extracted_curriculum = extract_file_text(curriculum_file)
        extracted_scheme = extract_file_text(scheme_file)
        extracted_notes = extract_file_text(notes_file)

        if not scheme_detail.strip() and not extracted_scheme and not extracted_curriculum:
            st.warning(f"Please upload your Curriculum Framework, Scheme of Work, or paste the objectives for {term_selected} {week_selected}.")
        else:
            with st.spinner(f"Parsing curriculum data for {term_selected} - {week_selected} and generating {lesson_count} lesson plan(s)..."):
                combined_docs = scheme_detail.strip()
                if extracted_curriculum:
                    combined_docs += f"\n\n--- CURRICULUM FRAMEWORK EXTRACT ---\n{extracted_curriculum[:8000]}"
                if extracted_scheme:
                    combined_docs += f"\n\n--- ANNUAL SCHEME OF WORK (LOCATE {term_selected.upper()} AND {week_selected.upper()}) ---\n{extracted_scheme[:12000]}"
                if extracted_notes:
                    combined_docs += f"\n\n--- REFERENCE NOTES / TEXTBOOK EXTRACT ---\n{extracted_notes[:5000]}"

                topic_instructions = []
                for k, v in contact_topics.items():
                    if v.strip():
                        topic_instructions.append(f"{k}: Use exact topic '{v.strip()}'")
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

                METADATA CONSTRAINTS:
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
                2. LESSON OBJECTIVES: Break down the scheme statements into active Bloom's Taxonomy verbs (Cognitive, Psychomotor, Affective), each retaining its exact curriculum framework code (e.g. [CHE1.1.1]).
                3. GUIDED PRACTICE: EVERY individual active learning step MUST explicitly begin with its matching curriculum framework code, e.g., "[CHE1.1.1] Students test unknown solutions using aqueous sodium hydroxide...".
                4. INSTRUCTIONAL RESOURCES & WORKING URL LINKS:
                   - The 'resources' field MUST include actual clickable text URLs:
                     * A direct YouTube search URL: "YouTube Video: [Video Title/Topic] - [https://www.youtube.com/results?search_query=](https://www.youtube.com/results?search_query=)[encoded+search+terms]"
                     * An interactive PhET simulation URL: "PhET Simulation: [Simulation Title] - [https://phet.colorado.edu/en/simulations/filter?subjects=](https://phet.colorado.edu/en/simulations/filter?subjects=)[topic]&type=html"
                     * Key concrete apparatus/reagents needed.
                5. LESSON SUMMARY NOTES: Provide 3-4 bullet points summarizing key concepts taught in this lesson (derived from the notes/framework). This will appear directly below Closure.
                6. CLOSURE (PLENARY): Must be exactly: "Concludes the lesson by giving a neat and tidy summary of the lesson.".
                7. HOD SECTION: The "hod_comment" field must be empty ("").
                8. Output must be strictly a JSON array with exactly {lesson_count} objects.

                JSON Schema keys per item:
                "school_name", "staff_name", "subject", "unit_topic", "lesson_topic", "date", "period", "week", "lesson_number", "sex", "duration", "class_name", "no_in_class", "objectives", "resources", "references", "prior_knowledge", "direct_teaching", "guided_practice", "evaluation", "closure", "summary_notes", "assignment", "hod_comment"
                """
                try:
                    client = genai.Client(api_key=api_key)
                    data = execute_generation_with_retry(client, prompt)
                    st.session_state.plans_docx = generate_docx_bytes(data, term_selected)
                    st.session_state.plans_filename = f"NGA_{term_selected.replace(' ', '_')}_{week_selected.replace(' ', '_')}_{subject}.docx"
                except Exception as e:
                    st.error(f"Generation error: {e}")

    if st.session_state.plans_docx:
        st.success("Inspection plans ready!")
        st.download_button(
            label="📥 Download Lesson Plans (.docx)",
            data=st.session_state.plans_docx,
            file_name=st.session_state.plans_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_plans"
        )

# =======================================================
# TAB 2: A* COMPREHENSIVE LESSON NOTE GENERATOR
# =======================================================
with tab_notes:
    st.subheader("🌟 Generate A* Cambridge/WAEC-Grade Lesson Notes")
    st.caption("Upload textbook chapters, revision guides, or syllabus extracts. Generates examiner-targeted notes with mandatory keywords, pitfall warnings, and graph figures.")

    an_c1, an_c2 = st.columns(2)
    with an_c1:
        an_subject = st.text_input("Subject", value="CHEMISTRY", key="an_sub")
        an_topic = st.text_input("Specific Topic", value="Rates of Reaction & Collision Theory", key="an_top")
        an_class = st.text_input("Class Target", value="Year 11 / Cambridge IGCSE", key="an_cls")
    with an_c2:
        graph_type = st.selectbox(
            "Auto-Embed Scientific Graph/Diagram Figure:",
            ["none", "Rate of Reaction (Volume vs Time)", "Heating/Cooling Curve (Phase Changes)", "Generic Trend Analysis"],
            index=1,
            key="an_graph"
        )
        textbook_file = st.file_uploader("Upload Textbook Chapter / Revision Note [.pdf, .docx, .txt]", type=["pdf", "docx", "txt"], key="an_file")

    an_extra_details = st.text_area(
        "Key Syllabus Codes or Focus Points (Optional):",
        placeholder="Focus on the effect of concentration, surface area, and catalyst on effective collisions.",
        height=80,
        key="an_extra"
    )

    if st.button("Generate A* Comprehensive Lesson Note", type="primary", key="btn_gen_notes"):
        extracted_textbook = extract_file_text(textbook_file)
        with st.spinner(f"Compiling A* comprehensive revision note for '{an_topic}'..."):
            note_prompt = f"""
            You are a Senior Principal Examiner preparing comprehensive, A* exam-targeted student lesson notes.
            SUBJECT: {an_subject}
            TOPIC: {an_topic}
            CLASS: {an_class}
            FOCUS: {an_extra_details}

            SOURCE TEXTBOOK / REVISION EXTRACT:
            {extracted_textbook[:12000] if extracted_textbook else "Use authoritative Cambridge IGCSE / WAEC curriculum specifications."}

            PEDAGOGICAL REQUIREMENTS FOR A* GRADE:
            1. EXAMINER MANDATORY KEYWORDS: Identify 4-6 technical terms examiners demand to see in written answers for full marks, providing concise definitions.
            2. DETAILED CORE SECTIONS: 3 to 4 distinct conceptual subheadings. Under each subheading, provide exhaustive, in-depth bullet points explaining the core mechanism, definitions, and equations. Do not skim or summarize vaguely.
            3. EXAMINER WARNING / COMMON PITFALLS: Provide 3-4 specific misconceptions where students lose marks (e.g., stating particles vibrate faster when heated in liquids, or confusing rate with yield).
            4. WORKED MODEL EXAMPLE: A step-by-step numerical or analytical question with a complete model answer showing how full marks are secured.
            5. Return strictly a single JSON object conforming to the schema.

            JSON Schema:
            {{
                "subject": "{an_subject}",
                "topic": "{an_topic}",
                "class_name": "{an_class}",
                "keywords": [
                    {{"term": "string", "meaning": "string"}}
                ],
                "sections": [
                    {{
                        "subheading": "string",
                        "points": ["string", "string", "string"]
                    }}
                ],
                "common_pitfalls": [
                    "string", "string"
                ],
                "worked_example": {{
                    "question": "string",
                    "solution_steps": ["string", "string"]
                }}
            }}
            """
            try:
                client = genai.Client(api_key=api_key)
                note_json = execute_generation_with_retry(client, note_prompt)
                st.session_state.note_docx = generate_astar_note_docx(note_json, graph_type)
                st.session_state.note_filename = f"NGA_AStar_Note_{an_topic.replace(' ', '_')}.docx"
            except Exception as e:
                st.error(f"Generation error: {e}")

    if st.session_state.note_docx:
        st.success(f"A* Lesson Note for '{an_topic}' generated successfully!")
        st.download_button(
            label="📥 Download A* Lesson Note (.docx)",
            data=st.session_state.note_docx,
            file_name=st.session_state.note_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_notes"
        )
