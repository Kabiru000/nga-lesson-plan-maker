import streamlit as st
import json
import io
import os
import re
import time
import fitz  # PyMuPDF
from PIL import Image
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

st.set_page_config(page_title="Noble Guide Academy Suite", layout="wide")

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

def extract_pdf_pages_and_render(uploaded_file, start_page: int, end_page: int):
    """
    Extracts text and renders full-page images so vector diagrams, charts, 
    and apparatus figures can be detected and cropped.
    """
    if uploaded_file is None:
        return "", []
    
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)
    
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    total_pages = len(doc)
    start_idx = max(0, start_page - 1)
    end_idx = min(total_pages, end_page)
    
    text_chunks = []
    rendered_pages = []

    # Render at 150 DPI for diagram clarity
    zoom = 150 / 72
    mat = fitz.Matrix(zoom, zoom)

    for page_idx in range(start_idx, end_idx):
        page = doc[page_idx]
        page_text = page.get_text()
        if page_text.strip():
            text_chunks.append(f"--- PAGE {page_idx + 1} ---\n{page_text}")
        
        pix = page.get_pixmap(matrix=mat)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        rendered_pages.append({
            "page_num": page_idx + 1,
            "image": img
        })

    return "\n\n".join(text_chunks), rendered_pages

def crop_diagram_from_box(page_img: Image.Image, box_2d: list) -> io.BytesIO:
    """
    Crops a diagram from normalized coordinates [ymin, xmin, ymax, xmax] on a 0-1000 scale.
    """
    w, h = page_img.size
    ymin, xmin, ymax, xmax = box_2d
    crop_box = (
        int((xmin / 1000) * w),
        int((ymin / 1000) * h),
        int((xmax / 1000) * w),
        int((ymax / 1000) * h)
    )
    cropped = page_img.crop(crop_box)
    buf = io.BytesIO()
    cropped.save(buf, format="PNG")
    buf.seek(0)
    return buf

def execute_generation_with_retry(client, prompt: str, contents=None):
    last_error = None
    for attempt in range(3):
        try:
            call_contents = contents if contents is not None else prompt
            response = client.models.generate_content(
                model="gemini-3.6-flash",
                contents=call_contents,
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

def generate_astar_note_docx(note_data: dict, cropped_diagrams: list) -> io.BytesIO:
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

    # Key Vocabulary
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

    # Core Scientific Sections
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

        # Attach topic-relevant diagram directly beneath the matching section
        if cropped_diagrams and sec_idx < len(cropped_diagrams):
            diag = cropped_diagrams[sec_idx]
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(2)
            doc.add_picture(diag["bytes"], width=Inches(4.8))
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_after = Pt(6)
            r_cap = p_cap.add_run(f"Figure: {diag['caption']} (Extracted from Textbook Page {diag['page']})")
            r_cap.font.size = Pt(8.5)
            r_cap.font.italic = True
            r_cap.font.color.rgb = RGBColor(90, 90, 90)

    # Pitfalls Warning Box
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

    # Worked Example
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
# MODE 1: A* COMPREHENSIVE LESSON NOTES (VISION DIAGRAMS)
# =======================================================
if app_mode == "🌟 A* Comprehensive Lesson Notes":
    st.title("🌟 A* Visual Lesson Note Generator")
    st.markdown("Generates A* revision notes and uses visual layout detection to extract relevant apparatus, graphs, and setup diagrams directly from the uploaded pages.")

    col1, col2 = st.columns(2)
    with col1:
        an_subject = st.text_input("Subject", value="CHEMISTRY")
        an_topic = st.text_input("Specific Topic", value="Acids, Bases and Salts")
        an_class = st.text_input("Class Target", value="Year 11 / Cambridge IGCSE")
    with col2:
        textbook_file = st.file_uploader("Upload Textbook / Revision PDF [.pdf]", type=["pdf"])
        p_col1, p_col2 = st.columns(2)
        with p_col1:
            start_p = st.number_input("Start Page Number", min_value=1, value=1, step=1)
        with p_col2:
            end_p = st.number_input("End Page Number", min_value=1, value=4, step=1)

    an_extra_details = st.text_area(
        "Target Curriculum Objectives / Codes (Optional):",
        placeholder="e.g. CHE1.1.1 Focus on apparatus setups, titration, or salt preparation.",
        height=70
    )

    if st.button("Generate Visual A* Lesson Note", type="primary"):
        if textbook_file is None:
            st.warning("Please upload the textbook PDF.")
        else:
            client = genai.Client(api_key=api_key)
            with st.spinner(f"Analyzing pages {start_p} to {end_p} and scanning for topic diagrams..."):
                extracted_text, rendered_pages = extract_pdf_pages_and_render(textbook_file, start_p, end_p)
                
                cropped_diagrams = []
                
                # Use Gemini Vision to detect and crop relevant diagrams
                for page_obj in rendered_pages:
                    detect_prompt = f"""
                    Examine this textbook page for the chemistry topic: '{an_topic}'.
                    If there is an apparatus diagram, experiment setup, reaction flowchart, or graph directly relevant to '{an_topic}', return its bounding box coordinates on a 0 to 1000 scale.
                    Do not detect decorative icons, company logos, or irrelevant figures.
                    If none found, return an empty array.

                    JSON Schema:
                    {{
                        "has_relevant_diagram": true,
                        "caption": "Concise caption explaining what this setup shows",
                        "box_2d": [ymin, xmin, ymax, xmax]
                    }}
                    """
                    try:
                        img_byte_arr = io.BytesIO()
                        page_obj["image"].save(img_byte_arr, format='JPEG', quality=85)
                        img_byte_arr.seek(0)
                        
                        vision_res = client.models.generate_content(
                            model="gemini-3.6-flash",
                            contents=[
                                types.Part.from_bytes(data=img_byte_arr.getvalue(), mime_type="image/jpeg"),
                                detect_prompt
                            ],
                            config=types.GenerateContentConfig(
                                temperature=0.1,
                                response_mime_type="application/json"
                            )
                        )
                        v_data = clean_json_response(vision_res.text)
                        if v_data.get("has_relevant_diagram") and v_data.get("box_2d"):
                            crop_buf = crop_diagram_from_box(page_obj["image"], v_data["box_2d"])
                            cropped_diagrams.append({
                                "bytes": crop_buf,
                                "caption": v_data.get("caption", "Apparatus Setup"),
                                "page": page_obj["page_num"]
                            })
                    except Exception:
                        continue

                st.info(f"📸 Detected and cropped {len(cropped_diagrams)} relevant diagram(s) from pages {start_p}–{end_p}.")

                # Generate grounded notes
                note_prompt = f"""
                You are a Senior Principal Examiner authoring an A* revision note.
                SUBJECT: {an_subject}
                TOPIC: {an_topic}
                CLASS: {an_class}
                CURRICULUM SPECIFICATIONS: {an_extra_details}

                SOURCE TEXTBOOK EXTRACT (PAGES {start_p} TO {end_p}):
                \"\"\"
                {extracted_text}
                \"\"\"

                CRITICAL FIDELITY REQUIREMENTS:
                1. STRICT GROUNDING: Derive definitions, chemical explanations, and procedures directly from the textbook extract.
                2. PROFESSIONAL EQUATION FORMATTING:
                   - Put balanced equations in the dedicated "equations" array under each section.
                   - Example: "CaCO3(s) + 2HCl(aq) -> CaCl2(aq) + H2O(l) + CO2(g)"
                3. UNCLUSTERED: Concise, well-spaced sentences in the "points" array. No raw markdown asterisks (**).
                4. EXAMINER MANDATORY VOCABULARY: 4 to 6 exact terms defined in the text.
                5. COMMON PITFALLS: 3 specific errors where students lose marks.
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
                    note_json = execute_generation_with_retry(client, note_prompt)
                    docx_file = generate_astar_note_docx(note_json, cropped_diagrams)

                    st.success(f"A* Lesson Note for '{an_topic}' compiled successfully!")
                    st.download_button(
                        label="📥 Download Visual A* Word Document (.docx)",
                        data=docx_file,
                        file_name=f"Visual_AStar_Note_{an_topic.replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"Generation error: {e}")

# =======================================================
# MODE 2: INSPECTORATE LESSON PLAN GENERATOR
# =======================================================
elif app_mode == "📋 Lesson Plan Generator":
    st.title("📋 Inspectorate Lesson Plan Generator")
    st.info("Lesson Plan generator is active. Use your standard workflow here.")
