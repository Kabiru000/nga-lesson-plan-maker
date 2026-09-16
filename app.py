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
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

st.set_page_config(page_title="Noble Guide Academy Suite", layout="wide")

# --- UI Styling Utilities ---
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

def add_styled_heading(doc, text: str, level: int = 1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.bold = True
    run.font.name = "Arial"
    if level == 1:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 51, 102) # Cambridge Navy
    elif level == 2:
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(40, 40, 40)
    return p

def format_chemical_text(paragraph, text: str):
    """
    Renders clean chemistry text, converting arrows and automatically
    subscripting numbers inside chemical formulas.
    """
    text = text.replace("->", " → ").replace("<=>", " ⇌ ")
    tokens = re.split(r'(\s+|[()+→⇌=Δ])', text)
    for token in tokens:
        if not token:
            continue
        # Check if token looks like a chemical formula with numbers (e.g., H2SO4, CaCO3)
        if re.search(r'[A-Z][a-z]?\d+', token):
            sub_parts = re.split(r'(\d+)', token)
            for sp in sub_parts:
                if sp.isdigit():
                    r = paragraph.add_run(sp)
                    r.font.subscript = True
                    r.font.name = "Calibri"
                    r.font.size = Pt(10)
                else:
                    r = paragraph.add_run(sp)
                    r.font.name = "Calibri"
                    r.font.size = Pt(10)
        else:
            r = paragraph.add_run(token)
            r.font.name = "Calibri"
            r.font.size = Pt(10)

# --- PDF Multimodal Extraction (Text + Embedded Images) ---
def extract_pdf_data_with_images(uploaded_file):
    """
    Extracts text and real diagram images from PDF using PyMuPDF.
    Filters out decorative icons/logos, returning real apparatus and figures.
    """
    if uploaded_file is None:
        return "", []
    
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)
    
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    extracted_text = []
    extracted_images = []

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        extracted_text.append(page.get_text())
        
        image_list = page.get_images(full=True)
        for img_info in image_list:
            xref = img_info[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            width = base_image["width"]
            height = base_image["height"]
            
            # Filter out tiny icons, buttons, banners (Keep only substantive diagrams)
            if width > 180 and height > 140 and (width * height > 35000):
                try:
                    img = Image.open(io.BytesIO(image_bytes))
                    # Convert CMYK/RGBA if needed to RGB for Word compatibility
                    if img.mode not in ("RGB", "L"):
                        img = img.convert("RGB")
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=90)
                    buf.seek(0)
                    extracted_images.append({
                        "bytes": buf,
                        "page": page_idx + 1,
                        "w": width,
                        "h": height
                    })
                except Exception:
                    continue

    return "\n".join(extracted_text), extracted_images

def clean_json_response(raw_text: str):
    text = raw_text.strip()
    if text.startswith("```json"):
        text = text[7:]
    elif text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    return json.loads(text.strip())

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

# --- Matplotlib Scientific Graph Engine ---
def generate_sample_topic_graph(graph_type: str) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=(6.2, 3.2), dpi=180)
    fig.patch.set_facecolor('#ffffff')
    ax.set_facecolor('#fdfdfd')

    if "rate" in graph_type.lower():
        t = np.linspace(0, 100, 200)
        c1 = 100 * (1 - np.exp(-0.05 * t))
        c2 = 100 * (1 - np.exp(-0.02 * t))
        ax.plot(t, c1, color='#004080', label='High Concentration / Catalyst (Steeper Initial Gradient)', linewidth=2.2)
        ax.plot(t, c2, color='#c0392b', linestyle='--', label='Lower Concentration / Room Temp', linewidth=2.0)
        ax.set_title("Reaction Kinetics: Gas Evolution vs Time", fontsize=10.5, fontweight='bold', pad=10)
        ax.set_xlabel("Time (s)", fontsize=9, fontweight='bold')
        ax.set_ylabel("Volume of Gas (cm³)", fontsize=9, fontweight='bold')
        ax.legend(fontsize=8, loc="lower right", framealpha=0.9)
        ax.grid(True, linestyle=':', alpha=0.5)
    elif "heating" in graph_type.lower() or "cooling" in graph_type.lower():
        time_pts = [0, 2, 5, 8, 11, 14]
        temps = [20, 80, 80, 120, 120, 140]
        ax.plot(time_pts, temps, color='#d35400', linewidth=2.4, marker='o', markersize=4)
        ax.set_title("Heating Curve: Constant Temperature Phase Transitions", fontsize=10.5, fontweight='bold', pad=10)
        ax.set_xlabel("Time of Heating (min)", fontsize=9, fontweight='bold')
        ax.set_ylabel("Temperature (°C)", fontsize=9, fontweight='bold')
        ax.annotate('Melting Point (Solid + Liquid)', xy=(3.5, 80), xytext=(1.5, 96),
                    arrowprops=dict(arrowstyle="->", color="black", lw=1.2))
        ax.annotate('Boiling Point (Liquid + Gas)', xy=(9.5, 120), xytext=(7.5, 134),
                    arrowprops=dict(arrowstyle="->", color="black", lw=1.2))
        ax.grid(True, linestyle=':', alpha=0.5)
    else:
        x = np.linspace(0, 10, 100)
        y = np.sin(x)
        ax.plot(x, y, color='#2c3e50', linewidth=2)
        ax.set_title("Scientific Quantitative Profile", fontsize=10.5, fontweight='bold', pad=10)
        ax.set_xlabel("Independent Variable", fontsize=9)
        ax.set_ylabel("Dependent Variable", fontsize=9)
        ax.grid(True, linestyle=':', alpha=0.5)

    plt.tight_layout()
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png')
    plt.close(fig)
    img_buf.seek(0)
    return img_buf

# --- DOCX Builder: Professional A* Lesson Note ---
def generate_astar_note_docx(note_data: dict, extracted_images: list, graph_type: str = "none") -> io.BytesIO:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    # 1. Publication Masthead
    p_mast = doc.add_paragraph()
    p_mast.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_mast.paragraph_format.space_before = Pt(0)
    p_mast.paragraph_format.space_after = Pt(2)
    
    r_inst = p_mast.add_run("NOBLE GUIDE ACADEMY, ABUJA\n")
    r_inst.bold = True
    r_inst.font.name = "Arial"
    r_inst.font.size = Pt(13)
    r_inst.font.color.rgb = RGBColor(0, 51, 102)

    r_banner = p_mast.add_run("A* TARGETED COMPREHENSIVE LEARNER RESOURCE\n")
    r_banner.bold = True
    r_banner.font.name = "Arial"
    r_banner.font.size = Pt(10)
    r_banner.font.color.rgb = RGBColor(110, 110, 110)

    # Topic Ribbon Table
    ribbon = doc.add_table(rows=1, cols=3)
    ribbon.style = "Table Grid"
    ribbon.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_w = (Inches(2.5), Inches(3.0), Inches(1.5))
    for c, w in zip(ribbon.rows[0].cells, col_w):
        c.width = w
        set_cell_margins(c, top=60, bottom=60, left=100, right=100)
        set_cell_shading(c, "003366") # Deep navy background

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

    doc.add_paragraph().paragraph_format.space_before = Pt(4)

    # 2. Examiner Mandatory Vocabulary Box
    t_voc = doc.add_table(rows=1, cols=1)
    t_voc.style = "Table Grid"
    t_voc.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_voc = t_voc.rows[0].cells[0]
    c_voc.width = Inches(7.0)
    set_cell_margins(c_voc, top=80, bottom=80, left=120, right=120)
    set_cell_shading(c_voc, "F4F7FA") # Premium light-slate tint

    p_vhead = c_voc.paragraphs[0]
    p_vhead.paragraph_format.space_after = Pt(4)
    r_vh = p_vhead.add_run("🔑 MANDATORY MARK-SCHEME TERMINOLOGY (EXAMINER CRITERIA)")
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
        r_term.font.size = Pt(9.5)
        format_chemical_text(p_k, kw.get('meaning', ''))

    # 3. Core Scientific Concepts & Formulas
    add_styled_heading(doc, "📘 IN-DEPTH CONCEPT MECHANISMS & PRINCIPLES", level=1)
    
    for sec_idx, section in enumerate(note_data.get("sections", [])):
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(6)
        p_sub.paragraph_format.space_after = Pt(2)
        r_sub = p_sub.add_run(f"§ {sec_idx + 1}. {section.get('subheading', '')}")
        r_sub.bold = True
        r_sub.font.name = "Arial"
        r_sub.font.size = Pt(10.5)
        r_sub.font.color.rgb = RGBColor(20, 20, 20)

        for pt in section.get("points", []):
            p_pt = doc.add_paragraph()
            p_pt.paragraph_format.left_indent = Inches(0.2)
            p_pt.paragraph_format.space_before = Pt(1)
            p_pt.paragraph_format.space_after = Pt(2)
            p_pt.add_run("• ").bold = True
            format_chemical_text(p_pt, pt)

        # Interleave uploaded diagrams into appropriate conceptual sections
        if extracted_images and sec_idx < len(extracted_images):
            img_item = extracted_images[sec_idx]
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(6)
            p_img.paragraph_format.space_after = Pt(2)
            doc.add_picture(img_item["bytes"], width=Inches(4.8))
            
            p_imgcap = doc.add_paragraph()
            p_imgcap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_imgcap.paragraph_format.space_after = Pt(6)
            r_cap = p_imgcap.add_run(f"Figure {sec_idx + 1}: Technical Diagram Extracted from Primary Course Material (Page {img_item['page']})")
            r_cap.font.name = "Arial"
            r_cap.font.size = Pt(8.5)
            r_cap.font.italic = True
            r_cap.font.color.rgb = RGBColor(80, 80, 80)

    # 4. Standard Plot Figure
    if graph_type != "none":
        add_styled_heading(doc, "📈 QUANTITATIVE TREND & GRAPH PROFILE INTERPRETATION", level=1)
        graph_buf = generate_sample_topic_graph(graph_type)
        doc.add_picture(graph_buf, width=Inches(5.5))
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_cap = p_cap.add_run(f"Figure: Graphical Analysis & Reaction Trajectory for {note_data.get('topic','')}")
        r_cap.font.name = "Arial"
        r_cap.font.size = Pt(8.5)
        r_cap.font.italic = True

    # 5. Examiner Misconceptions & Warning Box
    doc.add_paragraph().paragraph_format.space_before = Pt(4)
    t_warn = doc.add_table(rows=1, cols=1)
    t_warn.style = "Table Grid"
    t_warn.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_warn = t_warn.rows[0].cells[0]
    c_warn.width = Inches(7.0)
    set_cell_margins(c_warn, top=80, bottom=80, left=120, right=120)
    set_cell_shading(c_warn, "FDF3F2") # Muted Red/Warning tint

    p_whead = c_warn.paragraphs[0]
    p_whead.paragraph_format.space_after = Pt(4)
    r_wh = p_whead.add_run("⚠️ EXAMINER REPORT: COMMON PITFALLS & WHERE STUDENTS LOSE MARKS")
    r_wh.bold = True
    r_wh.font.name = "Arial"
    r_wh.font.size = Pt(10)
    r_wh.font.color.rgb = RGBColor(169, 50, 38)

    for pit in note_data.get("common_pitfalls", []):
        p_pit = c_warn.add_paragraph()
        p_pit.paragraph_format.space_before = Pt(1)
        p_pit.paragraph_format.space_after = Pt(2)
        p_pit.add_run("✖ ").bold = True
        format_chemical_text(p_pit, pit)

    # 6. Model Worked Calculation / Multi-Step Problem
    if note_data.get("worked_example"):
        doc.add_paragraph().paragraph_format.space_before = Pt(4)
        t_work = doc.add_table(rows=1, cols=1)
        t_work.style = "Table Grid"
        t_work.alignment = WD_TABLE_ALIGNMENT.CENTER
        c_work = t_work.rows[0].cells[0]
        c_work.width = Inches(7.0)
        set_cell_margins(c_work, top=80, bottom=80, left=120, right=120)
        set_cell_shading(c_work, "F9FBFD")

        p_workhead = c_work.paragraphs[0]
        p_workhead.paragraph_format.space_after = Pt(4)
        r_wkh = p_workhead.add_run("📝 A* STEP-BY-STEP WORKED MODEL SOLUTION")
        r_wkh.bold = True
        r_wkh.font.name = "Arial"
        r_wkh.font.size = Pt(10)
        r_wkh.font.color.rgb = RGBColor(0, 51, 102)

        p_wq = c_work.add_paragraph()
        r_wq_title = p_wq.add_run("Question: ")
        r_wq_title.bold = True
        format_chemical_text(p_wq, note_data['worked_example'].get('question', ''))

        for s_idx, stp in enumerate(note_data['worked_example'].get('solution_steps', [])):
            p_stp = c_work.add_paragraph()
            p_stp.paragraph_format.left_indent = Inches(0.2)
            r_st_title = p_stp.add_run(f"Step {s_idx + 1}: ")
            r_st_title.bold = True
            format_chemical_text(p_stp, stp)

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
# MODE 1: A* COMPREHENSIVE LESSON NOTES (VISUAL + EQUATIONS)
# =======================================================
if app_mode == "🌟 A* Comprehensive Lesson Notes":
    st.title("🌟 A* Publication-Grade Lesson Note Generator")
    st.markdown("Generates Cambridge/WAEC examination notes formatted with **native chemical equations**, **subscripts**, **examiner warning boxes**, and **diagrams extracted straight from uploaded materials**.")

    col1, col2 = st.columns(2)
    with col1:
        an_subject = st.text_input("Subject", value="CHEMISTRY")
        an_topic = st.text_input("Specific Topic", value="Acids, Bases and Salts")
        an_class = st.text_input("Class Target", value="Year 11 / Cambridge IGCSE")
    with col2:
        graph_type = st.selectbox(
            "Auto-Generate Kinetic/Thermal Graph:",
            ["none", "Rate of Reaction (Volume vs Time)", "Heating/Cooling Curve (Phase Changes)", "Generic Trend Analysis"],
            index=0
        )
        textbook_file = st.file_uploader("Upload Textbook Chapter / Revision Note [.pdf]", type=["pdf"])

    an_extra_details = st.text_area(
        "Specific Focus or Target Curriculum Codes:",
        placeholder="e.g. Focus on preparation of insoluble salts by precipitation, state symbols, and ionic equations.",
        height=80
    )

    if st.button("Generate Inspection-Grade A* Lesson Note", type="primary"):
        with st.spinner("Extracting diagrams, equations, and compiling A* revision master..."):
            extracted_text = ""
            extracted_images = []
            
            if textbook_file is not None:
                extracted_text, extracted_images = extract_pdf_data_with_images(textbook_file)
                st.info(f"📸 Detected and preserved {len(extracted_images)} high-resolution technical diagram(s) from your uploaded document.")

            note_prompt = f"""
            You are a Senior Cambridge & WAEC Chief Examiner authoring a definitive, publication-quality A* revision chapter.
            SUBJECT: {an_subject}
            TOPIC: {an_topic}
            CLASS: {an_class}
            FOCUS: {an_extra_details}

            SOURCE MATERIAL TEXT:
            {extracted_text[:12000] if extracted_text else "Use authoritative Cambridge IGCSE / WAEC examination syllabus standards."}

            STRICT COMPILATION STANDARDS:
            1. PROFESSIONAL CHEMICAL EQUATIONS & NOTATION:
               - Write all chemical formulas accurately with proper notation (e.g., CaCO3, H2SO4, Cu(NO3)2).
               - Always include state symbols: (s), (l), (g), (aq).
               - Explicitly write full ionic equations where relevant (e.g. Ba2+(aq) + SO4 2-(aq) -> BaSO4(s)).
               - Write reversible reactions using <=> and standard reactions with ->.
            2. EXAMINER MANDATORY VOCABULARY: 4 to 6 technical terms that examiners strictly mark for (e.g., 'proton donor', 'lattice', 'spectator ions').
            3. IN-DEPTH CORE BREAKDOWN: 3 to 4 substantial sections. Provide rigorous, sentence-length scientific explanations with equations. Do not provide superficial summaries.
            4. COMMON PITFALLS: 3 to 4 specific misconceptions where B-grade students fail to secure A* marks.
            5. WORKED MODEL SOLUTION: A multi-step calculation or synthesis preparation method showing full marks.
            6. Return strictly a JSON object.

            JSON Schema:
            {{
                "subject": "{an_subject}",
                "topic": "{an_topic}",
                "class_name": "{an_class}",
                "keywords": [{{"term": "string", "meaning": "string"}}],
                "sections": [{{"subheading": "string", "points": ["string", "string", "string"]}}],
                "common_pitfalls": ["string", "string"],
                "worked_example": {{"question": "string", "solution_steps": ["string", "string"]}}
            }}
            """
            try:
                client = genai.Client(api_key=api_key)
                note_json = execute_generation_with_retry(client, note_prompt)
                docx_file = generate_astar_note_docx(note_json, extracted_images, graph_type)

                st.success(f"A* Lesson Note for '{an_topic}' successfully compiled!")
                st.download_button(
                    label="📥 Download Publication-Grade A* Word Document (.docx)",
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
        st.info("Lesson Plan generator is active. Use your existing workflow here.")
