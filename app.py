import streamlit as st
import json
import io
import os
import time
from pypdf import PdfReader
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

st.set_page_config(page_title="Noble Guide Academy Lesson Plan Generator", layout="wide")

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
            ("Instructional/Teaching Resources:", plan.get('resources', ''), "hyphen"),
            ("Reference(s):", plan.get('references', ''), "none"),
            ("Prior Knowledge and Connection:", plan.get('prior_knowledge', ''), "none"),
            ("Direct Teaching:", plan.get('direct_teaching', ''), "diamond"),
            ("Guided Practice (Students’ Active Learning):", plan.get('guided_practice', ''), "diamond"),
            ("Lesson Summary Notes:", plan.get('summary_notes', ''), "hyphen"),
            ("Evaluation:", plan.get('evaluation', ''), "hyphen"),
            ("Closure (Plenary):", plan.get('closure', 'Concludes the lesson by giving a neat and tidy summary of the lesson.'), "none"),
            ("Assignment:", plan.get('assignment', ''), "hyphen"),
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
            return json.loads(response.text)
        except Exception as err:
            err_str = str(err)
            last_error = err
            if any(k in err_str for k in ["503", "UNAVAILABLE"]):
                time.sleep(2 * (attempt + 1))
                continue
            raise err
    raise last_error

query_params = st.query_params
saved_key = query_params.get("k", "")

with st.sidebar:
    st.header("🔑 API Configuration")
    user_key = st.text_input(
        "Gemini API Key:",
        value=saved_key,
        type="password",
        help="Paste your free API key from aistudio.google.com"
    )

    if st.button("💾 Save Key to this Device"):
        if user_key.strip():
            st.query_params["k"] = user_key.strip()
            st.success("Key saved! Bookmark this URL.")
        else:
            st.query_params.clear()
            st.info("Key cleared.")

    st.markdown("---")
    st.markdown(
        "**Get a Free Key:**\n"
        "1. Open [aistudio.google.com](https://aistudio.google.com)\n"
        "2. Click **Get API key**\n"
        "3. Paste here and Save."
    )

api_key = user_key.strip()

st.title("📚 Noble Guide Academy Lesson Plan Generator")
st.markdown("Annual Session Curriculum Management: 3 Terms × 8 Weeks × 1–5 Contacts per Subject")

if not api_key:
    st.warning("👈 Please enter your Gemini API key in the left sidebar to start generating plans.")
    st.stop()

# --- SESSION SCHEDULE SELECTORS ---
st.subheader("🗓️ Session & Term Scheduling")
s_col1, s_col2, s_col3 = st.columns(3)
with s_col1:
    term_selected = st.selectbox("Select Term", ["Term 1 (First Term)", "Term 2 (Second Term)", "Term 3 (Third Term)"])
with s_col2:
    week_selected = st.selectbox("Select Week", [f"Week {i}" for i in range(1, 9)])
with s_col3:
    lesson_count = st.number_input("Contacts/Lessons for this Subject this Week", min_value=1, max_value=5, value=3, step=1)

# --- CLASS & TEACHER METADATA ---
st.subheader("📋 Teacher & Class Information")
c1, c2 = st.columns(2)
with c1:
    staff_name = st.text_input("Teacher's Full Name", value="AMINU KABIRU")
    subject = st.text_input("Subject", value="CHEMISTRY")
    class_name = st.text_input("Class", value="Year 11")
with c2:
    date_schedule = st.text_input("Date(s) for this Week", value="7th & 8th June, 2026")
    duration = st.text_input("Duration of Lesson", value="50minutes")
    no_in_class = st.text_input("Number in Class", value="20")

c3, c4 = st.columns(2)
with c3:
    unit_topic = st.text_input("Unit Topic for this Week", value="Chemical and investigations")
with c4:
    textbooks = st.text_input("Reference Textbooks", value="New School Chemistry and Cambridge Chemistry Syllabus IGCSE Course Book")

# Dynamic Lesson Topic Configuration (One distinct input per contact)
st.subheader(f"📖 Specific Lesson Topics for {lesson_count} Contact(s)")
st.caption("Enter the exact topic for each lesson. If left blank, the AI will derive it from your curriculum/scheme.")

contact_topics = {}
topic_cols = st.columns(int(lesson_count))
for i in range(int(lesson_count)):
    with topic_cols[i]:
        contact_topics[f"Lesson {i+1}"] = st.text_input(
            f"Lesson {i+1} Topic:",
            placeholder=f"e.g., Sub-topic {i+1}",
            key=f"topic_contact_{i+1}"
        )

# 3 Dedicated File Uploaders
st.subheader("📎 Curriculum, Scheme & Note Uploads (Optional)")
u_col1, u_col2, u_col3 = st.columns(3)
with u_col1:
    curriculum_file = st.file_uploader("1. Curriculum Framework [.pdf, .docx, .txt]", type=["pdf", "docx", "txt"])
with u_col2:
    scheme_file = st.file_uploader("2. Annual Scheme of Work [.pdf, .docx, .txt]", type=["pdf", "docx", "txt"])
with u_col3:
    notes_file = st.file_uploader("3. Lesson Notes / Textbook [.pdf, .docx, .txt]", type=["pdf", "docx", "txt"])

scheme_detail = st.text_area(
    f"Curriculum Objectives & Codes for {term_selected}, {week_selected} (Optional: paste here if not uploading files):",
    height=100,
    placeholder="CHE1.1.1 Identify Cations and Anions in a solution.\nCHE1.1.2 Test for aqueous cations using sodium hydroxide."
)

if st.button("Generate Inspection Plans", type="primary"):
    extracted_curriculum = extract_file_text(curriculum_file)
    extracted_scheme = extract_file_text(scheme_file)
    extracted_notes = extract_file_text(notes_file)

    if not scheme_detail.strip() and not extracted_scheme and not extracted_curriculum:
        st.warning(f"Please upload your Curriculum Framework, Scheme of Work, or paste the objectives for {term_selected} {week_selected}.")
    else:
        with st.spinner(f"Parsing curriculum data for {term_selected} - {week_selected} and generating {lesson_count} lesson plan(s)..."):
            combined_docs = scheme_detail.strip()
            if extracted_curriculum:
                combined_docs += f"\n\n--- CURRICULUM FRAMEWORK EXTRACT ---\n{extracted_curriculum[:6000]}"
            if extracted_scheme:
                combined_docs += f"\n\n--- ANNUAL SCHEME OF WORK (LOCATE {term_selected.upper()} AND {week_selected.upper()}) ---\n{extracted_scheme[:10000]}"
            if extracted_notes:
                combined_docs += f"\n\n--- REFERENCE NOTES / TEXTBOOK EXTRACT ---\n{extracted_notes[:5000]}"

            # Format the specific user-provided contact topics
            topic_instructions = []
            for k, v in contact_topics.items():
                if v.strip():
                    topic_instructions.append(f"{k}: Use exact topic '{v.strip()}'")
                else:
                    topic_instructions.append(f"{k}: Derive appropriate specific sub-topic from the scheme")
            topic_summary = "\n".join(topic_instructions)

            prompt = f"""
            You are an expert Inspectorate Curriculum Specialist for Noble Guide Academy, Abuja.
            The school operates an annual session structured as 3 terms, each having 8 teaching weeks, with subjects having between 1 to 5 contacts/lessons per week.

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

            PEDAGOGICAL REQUIREMENTS:
            1. LESSON TOPIC FIDELITY: For each lesson, set the "lesson_topic" field to the corresponding topic specified above. If the teacher did not supply an exact topic, derive a precise, granular sub-topic for that period.
            2. TARGETED EXTRACTION: Specifically isolate and teach the curriculum codes and concepts assigned to {term_selected} and {week_selected}.
            3. BLOOM'S TAXONOMY DERIVATION: Break down syllabus statements into measurable objectives starting with active Bloom's verbs across Cognitive, Psychomotor, and Affective domains. Each objective must display its syllabus code (e.g., [CHE1.1.1]).
            4. GUIDED PRACTICE: EVERY individual active learning step MUST begin with its corresponding framework code, e.g., "[CHE1.1.1] Students examine unknown salt solutions in test tubes...".
            5. LESSON SUMMARY NOTES: Generate 3-4 bullet points summarizing the core teaching content for this lesson (derived from the scheme/notes/framework).
            6. HOD SECTION: The "hod_comment" field must be empty ("").
            7. RESOURCES:
               - Specific YouTube search video recommendation: "Video: [Title] - [Channel] (YouTube)"
               - Simulation/lab apparatus: "Simulation: PhET Interactive Simulations - [Topic]" or lab reagents.
            8. PRIOR KNOWLEDGE: format "The students are familiar with...".
            9. CLOSURE: "Concludes the lesson by giving a neat and tidy summary of the lesson.".
            10. Output must be strictly a JSON array with exactly {lesson_count} objects (Lesson 1 to Lesson {lesson_count}).

            JSON Schema keys per item:
            "school_name", "staff_name", "subject", "unit_topic", "lesson_topic", "date", "period", "week", "lesson_number", "sex", "duration", "class_name", "no_in_class", "objectives", "resources", "references", "prior_knowledge", "direct_teaching", "guided_practice", "summary_notes", "evaluation", "closure", "assignment", "hod_comment"
            """
            try:
                client = genai.Client(api_key=api_key)
                data = execute_generation_with_retry(client, prompt)
                file_data = generate_docx_bytes(data, term_selected)

                st.success(f"Generated {len(data)} inspection lesson plan(s) for {term_selected}, {week_selected}!")
                st.download_button(
                    label=f"📥 Download Word Document (.docx)",
                    data=file_data,
                    file_name=f"NGA_{term_selected.replace(' ', '_')}_{week_selected.replace(' ', '_')}_{subject}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Generation error: {e}")
