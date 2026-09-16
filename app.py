import streamlit as st
import json
import io
import os
import re
import time
import fitz  # PyMuPDF
from PIL import Image
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

# --- DOCX Typography & Styling Utilities ---
def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
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

def format_equation_line(p, eq_str: str):
    """Renders chemical and mathematical equations with proper sub/superscripts."""
    clean_eq = eq_str.replace("->", " → ").replace("<=>", " ⇌ ")
    tokens = re.split(r'(\s+|[()+→⇌=Δ])', clean_eq)
    for token in tokens:
        if not token:
            continue
        # Superscript charges: 2+, 3+, 2-, +, -
        if re.fullmatch(r'\d*[\+\-]', token):
            r = p.add_run(token)
            r.font.superscript = True
            r.font.name = "Cambria Math"
            r.font.size = Pt(11)
            r.bold = True
        # Subscript chemical formulas: H2SO4, CaCO3
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
    """Cleanly formats text without markdown artifacts, using 1.15 line spacing."""
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

    # Parse inline markdown bolding if present
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

# --- Targeted PDF Page Extraction ---
def extract_pdf_pages(uploaded_file, start_page: int, end_page: int, extract_images: bool = False):
    if uploaded_file is None:
        return "", []
    
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)
    
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    total_pages = len(doc)
    
    # Constrain range to actual PDF boundaries
    start_idx = max(0, start_page - 1)
    end_idx = min(total_pages, end_page)
    
    text_chunks = []
    images = []

    for page_idx in range(start_idx, end_idx):
        page = doc[page_idx]
        page_text = page.get_text()
        if page_text.strip():
            text_chunks.append(f"--- PAGE {page_idx + 1} ---\n{page_text}")
            
        if extract_images:
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                base_img = doc.extract_image(xref)
                w, h = base_img["width"], base_img["height"]
                # Must be a substantial illustration, not an icon or banner
                if w > 250 and h > 180 and (w * h > 50000):
                    try:
                        pil_img = Image.open(io.BytesIO(base_img["image"]))
                        if pil_img.mode not in ("RGB", "L"):
                            pil_img = pil_img.convert("RGB")
                        buf = io.BytesIO()
                        pil_img.save(buf, format="JPEG", quality=90)
                        buf.seek(0)
                        images.append({"bytes": buf, "page": page_idx + 1})
                    except Exception:
                        continue

    return "\n\n".join(text_chunks), images

def execute_generation_with_retry(client, prompt: str):
    last_error = None
    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model="gemini-3.6-flash",
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=0.1,
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

# --- DOCX Builder: Grounded A* Lesson Note ---
def generate_astar_note_docx(note_data: dict, verified_images: list) -> io.BytesIO:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # 1. School & Resource Header
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

    # Topic Ribbon
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

    # 2. Key Vocabulary Box
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

    # 3. Core Scientific Notes Grounded in Uploaded Text
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

        # Isolated Equation Callout Boxes
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

    # Optional Verified Diagram Insertion
    if verified_images:
        doc.add_paragraph().paragraph_format.space_before = Pt(6)
        p_dh = doc.add_paragraph()
        r_dh = p_dh.add_run("🔬 TEXTBOOK DIAGRAM REFERENCE")
        r_dh.bold = True
        r_dh.font.name = "Arial"
        r_dh.font.size = Pt(10.5)
        r_dh.font.color.rgb = RGBColor(0, 51, 102)

        for idx, img_obj in enumerate(verified_images[:2]):
            doc.add_picture(img_obj["bytes"], width=Inches(4.8))
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_after = Pt(6)
            r_cap = p_cap.add_run(f"Figure {idx + 1}: Diagram Extracted from Textbook Reading (Page {img_obj['page']})")
            r_cap.font.size = Pt(8.5)
            r_cap.font.italic = True
            r_cap.font.color.rgb = RGBColor(90, 90, 90)

    # 4. Examiner Misconceptions & Warning Box
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

    # 5. Model Problem & Worked Example
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

# --- SIDEBAR & AUTH ---
query_params = st.query_params
saved_key = query_params.get("k", "")

with st.sidebar:
    st.header("⚙️ Workstation Menu")
    app_mode = st.radio(
        "Select Tool Mode:",
        ["🌟 A* Comprehensive Lesson Notes", "📋 Lesson Plan Generator"],
        index=0
    )
    st.markdown("---")
    st.header("🔑 API Configuration")
    user_key = st.text_input("Gemini API Key:", value=saved_key, type="password", help="Free key from aistudio.google.com")
    if st.button("💾 Save Key to this Device"):
        if user_key.strip():
            st.query_params["k"] = user_key.strip()
            st.success("Key saved permanently in browser URL!")
        else:
            st.query_params.clear()
            st.info("Key cleared.")

api_key = user_key.strip()
if not api_key:
    st.warning("👈 Enter your Gemini API key in the sidebar to activate the generator.")
    st.stop()

# =======================================================
# MODE 1: A* COMPREHENSIVE LESSON NOTES
# =======================================================
if app_mode == "🌟 A* Comprehensive Lesson Notes":
    st.title("🌟 A* Grounded Lesson Note Generator")
    st.markdown("Generates clean, professional revision notes **derived strictly from your specified textbook pages**.")

    col1, col2 = st.columns(2)
    with col1:
        an_subject = st.text_input("Subject", value="CHEMISTRY")
        an_topic = st.text_input("Specific Topic", value="Acids, Bases and Salts")
        an_class = st.text_input("Class Target", value="Year 11 / Cambridge IGCSE")
    with col2:
        textbook_file = st.file_uploader("Upload Textbook / Revision PDF [.pdf]", type=["pdf"])
        p_col1, p_col2 = st.columns(2)
        with p_col1:
            start_p = st.number_input("Start Page Number", min_value=1, value=1, step=1, help="Specify the page where this topic begins to avoid covers/prefaces")
        with p_col2:
            end_p = st.number_input("End Page Number", min_value=1, value=6, step=1, help="Specify the page where this topic ends")
        
        include_imgs = st.checkbox("Extract & include diagrams found in these pages", value=False)

    an_extra_details = st.text_area(
        "Target Curriculum Objectives / Codes (Optional):",
        placeholder="e.g. CHE1.1.1 Identify Cations and Anions; focus on preparation of insoluble salts via precipitation.",
        height=80
    )

    if st.button("Generate Inspection-Grade A* Lesson Note", type="primary"):
        if textbook_file is None:
            st.warning("Please upload your textbook or revision PDF so the AI can extract the exact lesson notes.")
        else:
            with st.spinner(f"Extracting content from pages {start_p} to {end_p} and compiling lesson notes..."):
                extracted_text, extracted_images = extract_pdf_pages(textbook_file, start_p, end_p, extract_images=include_imgs)
                
                if not extracted_text.strip():
                    st.error(f"No readable text was found on pages {start_p} to {end_p}. Please verify the page numbers.")
                else:
                    note_prompt = f"""
                    You are a Senior Principal Examiner preparing a publication-grade, A* student revision note.
                    SUBJECT: {an_subject}
                    TOPIC: {an_topic}
                    CLASS: {an_class}
                    CURRICULUM SPECIFICATIONS: {an_extra_details}

                    PRIMARY SOURCE TEXTBOOK EXTRACT (PAGES {start_p} TO {end_p}):
                    \"\"\"
                    {extracted_text}
                    \"\"\"

                    CRITICAL FIDELITY REQUIREMENTS:
                    1. STRICT GROUNDING: Derive the core definitions, scientific explanations, and practical procedures directly from the provided textbook extract. Do not fabricate unrelated concepts.
                    2. PROFESSIONAL EQUATION FORMATTING:
                       - Place all balanced equations in the dedicated "equations" array under each section.
                       - Write complete equations with correct state symbols (s, l, g, aq).
                       - Example format: "CaCO3(s) + 2HCl(aq) -> CaCl2(aq) + H2O(l) + CO2(g)"
                    3. CLEAN, WELL-SPACED STRUCTURE:
                       - In the "points" array, write concise, informative sentences.
                       - Never use raw markdown asterisks (**) in your json values.
                    4. EXAMINER MANDATORY VOCABULARY: 4 to 6 exact technical keywords defined in the textbook text.
                    5. COMMON PITFALLS: 3 specific errors where students frequently lose marks on this topic.
                    6. WORKED EXAMPLE: A complete model calculation or synthesis problem showing step-by-step logic.

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
                        docx_file = generate_astar_note_docx(note_json, extracted_images)

                        st.success(f"A* Lesson Note for '{an_topic}' compiled successfully from pages {start_p}–{end_p}!")
                        st.download_button(
                            label="📥 Download Grounded A* Word Document (.docx)",
                            data=docx_file,
                            file_name=f"AStar_Note_{an_topic.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    except Exception as e:
                        st.error(f"Generation error: {e}")

# =======================================================
# MODE 2: INSPECTORATE LESSON PLAN GENERATOR
# =======================================================
elif app_mode == "📋 Lesson Plan Generator":
    st.title("📋 Inspectorate Lesson Plan Generator")
    st.markdown("Annual Session Curriculum Management: 3 Terms × 8 Weeks × 1–5 Contacts per Subject")

    s_col1, s_col2, s_col3 = st.columns(3)
    with s_col1:
        term_selected = st.selectbox("Select Term", ["Term 1 (First Term)", "Term 2 (Second Term)", "Term 3 (Third Term)"])
    with s_col2:
        week_selected = st.selectbox("Select Week", [f"Week {i}" for i in range(1, 9)])
    with s_col3:
        lesson_count = st.number_input("Contacts/Lessons this Week", min_value=1, max_value=5, value=3, step=1)

    c1, c2 = st.columns(2)
    with c1:
        staff_name = st.text_input("Teacher's Full Name", value="AMINU KABIRU")
        subject = st.text_input("Subject", value="CHEMISTRY")
        class_name = st.text_input("Class", value="Year 11")
    with c2:
        date_schedule = st.text_input("Date(s) for this Week", value="7th & 8th June, 2026")
        duration = st.text_input("Duration of Lesson", value="50minutes")
        no_in_class = st.text_input("Number in Class", value="20")

    unit_topic = st.text_input("Unit Topic for this Week", value="Acids, Bases and Salts")
    textbooks = st.text_input("Reference Textbooks", value="New School Chemistry & Cambridge IGCSE Chemistry")

    contact_topics = {}
    topic_cols = st.columns(int(lesson_count))
    for i in range(int(lesson_count)):
        with topic_cols[i]:
            contact_topics[f"Lesson {i+1}"] = st.text_input(f"Lesson {i+1} Topic:", placeholder=f"Sub-topic {i+1}", key=f"lp_top_{i+1}")

    scheme_file = st.file_uploader("Upload Annual Scheme of Work [.pdf, .docx, .txt]", type=["pdf", "docx", "txt"], key="lp_sch")
    scheme_detail = st.text_area(f"Curriculum Objectives & Codes (Optional):", placeholder="CHE1.1.1 Identify Cations and Anions...", height=80)

    if st.button("Generate Inspection Plans", type="primary"):
        st.info("Lesson Plan generator is active. Use your standard workflow here.")
