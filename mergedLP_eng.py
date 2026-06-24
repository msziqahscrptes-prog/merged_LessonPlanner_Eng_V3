import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- 1. APP CONFIGURATION ---
st.set_page_config(page_title="PEDATI Lesson Planner", layout="wide")
st.title("🎓 PTES PEDATI FLOW LESSON PLANNER")

# --- MAIN PAGE CONFIGURATION & USER API KEY BAR (AT THE VERY TOP) ---
user_api_key = st.text_input(
    "🔑 ENTER YOUR GEMINI API KEY:", 
    type="password", 
    help="Get your API key from Google AI Studio using your Gmail account."
)

# Helper function to validate and load models dynamically
def get_working_model(api_key):
    try:
        genai.configure(api_key=api_key)
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                return m.name
    except Exception as e:
        st.error(f"INVALID API KEY OR CONNECTION ERROR: {str(e)}")
        return None
    return "models/gemini-1.5-flash"

selected_model_name = None
if user_api_key:
    selected_model_name = get_working_model(user_api_key)
    if selected_model_name:
        st.info(f"SYSTEM CONNECTED VIA YOUR API KEY. ACTIVE MODEL: {selected_model_name.upper()}")
else:
    st.warning("⚠️ PLEASE ENTER YOUR PERSONAL GEMINI API KEY ABOVE TO START.")


# --- 2. RESILIENT AI PROMPT GENERATION ENGINE ---
def generate_pedati_plan(topic, syllabus, extra_context, api_key, model_name):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)
    
    prompt = f"""
    Topic: {topic}. Syllabus Code: {syllabus}. Context: {extra_context}.
    Generate a professional lesson plan in English based strictly on the PEDATI framework.
    
    CRITICAL TEXT FORMATTING RULES:
    1. DO NOT use double asterisks (**) anywhere in the output.
    2. Ensure every single section title and heading is in full CAPITAL LETTERS.
    3. DO NOT use bullets or dots. Use a numbering system (1, 2, 3...) for item lists.
    
    Use the following EXACT structural markers for the output text format:
    
    SECTION: TOPIC
    {topic.upper()}
    
    SECTION: LESSON OBJECTIVES
    [4 points using numbers 1 to 4]
    
    SECTION: LESSON OUTCOMES
    [4 points using numbers 1 to 4]
    
    SECTION: SUCCESS CRITERIA
    [4 points using numbers 1 to 4]
    
    SECTION: PREREQUISITE
    [1 point using number 1]
    
    SECTION: KEYWORDS
    [6 items using numbers 1 to 6]
    
    SECTION: HOTS
    [4 main domains from Bloom's Taxonomy using numbers 1 to 4]
    
    SECTION: DIGITAL CITIZENSHIP
    [4 points on ethical tech use/Chromebooks/Canva/YouTube using numbers 1 to 4]

    SECTION: OPENING LESSON CONTENT
    [Hook activity and transition plan]

    SECTION: DIFFERENTIATION STRATEGIES (GREEN)
    1. HA (Higher Achiever): [1 challenging activity explained as a numbered item]

    SECTION: DIFFERENTIATION STRATEGIES (YELLOW)
    1. MA (Medium Achiever): [1 core activity explained as a numbered item]

    SECTION: DIFFERENTIATION STRATEGIES (RED)
    1. LA (Lower Achiever): [1 scaffolded activity explained as a numbered item]

    SECTION: BLENDED LEARNING ACTIVITY ONE (15 MINS)
    1. Activity 1: [Descriptions]
    2. Teacher Preparation: [Step-by-step before lesson as numbered sub-items]
    3. Objectives: [3 points using numbers]
    4. Student Tasks: [Step-by-step details as numbered sub-items]

    SECTION: BLENDED LEARNING ACTIVITY TWO (15 MINS)
    1. Activity 2: [Descriptions]
    2. Teacher Preparation: [Step-by-step before lesson as numbered sub-items]
    3. Objectives: [3 points using numbers]
    4. Student Tasks: [Step-by-step details as numbered sub-items]

    SECTION: PEDATI FLOW GRID
    BLOCK_START: P: PREPARATION (LEARN)
    LECTURER: [Actionable steps aligned with the topic]
    STUDENTS: [Actionable tasks/chromebook work aligned with the topic]
    BLOCK_END
    
    BLOCK_START: E: ENGAGE (EXPLORE)
    LECTURER: [Actionable steps aligned with the topic]
    STUDENTS: [Actionable tasks/chromebook work aligned with the topic]
    BLOCK_END

    BLOCK_START: D.A: DELIVER AND APPLY
    LECTURER: [Actionable steps aligned with the topic]
    STUDENTS: [Actionable tasks/chromebook work aligned with the topic]
    BLOCK_END

    BLOCK_START: T.I: TEST AND EVALUATE
    LECTURER: [Actionable steps aligned with the topic]
    STUDENTS: [Actionable tasks/chromebook work aligned with the topic]
    BLOCK_END
    
    SECTION: PLENARY (EXIT TICKET)
    [2-3 minute closing activity]

    SECTION: HOMEWORK
    [Task assigned based on topic]

    SECTION: SUGGESTED WAY FORWARD TASK
    [Hook activity and transition plan for next lesson]
    """
    try:
        response = model.generate_content(prompt)
        if response.candidates and response.candidates[0].content.parts:
            # Absolute hard cleaning filter for asterisk protection
            return response.text.replace("**", "")
        else:
            return "⚠️ SYSTEM ERROR: THE AI RETURNED AN EMPTY RESPONSE. PLEASE TRY AGAIN."
    except Exception as e:
        return f"SYSTEM ERROR: {str(e)}"


# --- 3. WORD DOCUMENT GENERATION ENGINE (14PT FONT & 1.0 SPACING) ---
def create_word_export(topic, syllabus, text):
    doc = Document()
    
    # Baseline Page Document Styles Override
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(14)  # Force 14-Point Font Size Globally
    
    # Set Paragraph Format spacing rules globally (1 Paragraph / Single Spacing)
    p_format = style.paragraph_format
    p_format.line_spacing = 1.0
    p_format.space_after = Pt(12)  # Balanced block spacing

    # Document Header Title
    title_p = doc.add_paragraph()
    run_t = title_p.add_run(f'PTES PEDATI LESSON PLAN: {topic.upper()}')
    run_t.bold = True
    run_t.font.size = Pt(18)

    # Admin Info Table Setup
    admin_table = doc.add_table(rows=3, cols=4)
    admin_table.style = 'Table Grid'
    labels = [["WEEK NO:", "DATE:"], ["CLASS SIZE:", "DAY:"], ["VENUE:", "DURATION:"]]
    for r in range(3):
        admin_table.cell(r, 0).paragraphs[0].add_run(labels[r][0]).bold = True
        admin_table.cell(r, 2).paragraphs[0].add_run(labels[r][1]).bold = True
    doc.add_paragraph()

    # Split document cleanly based on Section Markers
    sections = text.split('SECTION:')
 
    for section in sections:
        if not section.strip(): 
            continue
            
        lines = section.strip().split('\n')
        title = lines[0].strip().upper().replace("**", "")
        body_content = "\n".join(lines[1:]).strip()

        if "PEDATI FLOW GRID" in title:
            h = doc.add_paragraph()
            r_h = h.add_run("P.E.D.A.T.I FLOW BREAKDOWN")
            r_h.bold = True
            
            blocks = body_content.split("BLOCK_START:")
            for block in blocks:
                if not block.strip(): 
                    continue
                block_data = block.split("BLOCK_END")[0].strip().split('\n')
                
                heading_title = block_data[0].strip().upper().replace("**", "")
                lecturer_text = ""
                students_text = ""
                
                for line in block_data:
                    if line.upper().startswith("LECTURER:"):
                        lecturer_text = line.split(":", 1)[1].strip().replace("**", "")
                    elif line.upper().startswith("STUDENTS:"):
                        students_text = line.split(":", 1)[1].strip().replace("**", "")
                
                bp = doc.add_paragraph()
                brun = bp.add_run(heading_title)
                brun.bold = True
                
                table = doc.add_table(rows=2, cols=2)
                table.style = 'Table Grid'
                
                for row in table.rows:
                    row.cells[0].width = Inches(3.25)
                    row.cells[1].width = Inches(3.25)
                
                hdr_cells = table.rows[0].cells
                hdr_cells[0].paragraphs[0].add_run("LECTURER").bold = True
                hdr_cells[1].paragraphs[0].add_run("STUDENTS").bold = True
                
                # Apply 14pt custom text to cells safely
                c0_p = table.rows[1].cells[0].paragraphs[0]
                c0_p.paragraph_format.line_spacing = 1.0
                c1_p = table.rows[1].cells[1].paragraphs[0]
                c1_p.paragraph_format.line_spacing = 1.0
                
                c0_p.add_run(lecturer_text)
                c1_p.add_run(students_text)
                doc.add_paragraph()
        else:
            content = body_content.replace("**", "") 
            
            # Formally append sections using CAPITAL LETTERS titles
            hp = doc.add_paragraph()
            hrun = hp.add_run(title)
            hrun.bold = True
            
            table = doc.add_table(rows=1, cols=1)
            table.style = 'Table Grid'
            cell_p = table.cell(0, 0).paragraphs[0]
            cell_p.paragraph_format.line_spacing = 1.0
            cell_p.add_run(content)
            doc.add_paragraph()
     
    # Final Executive HOD Sign-off Grid
    doc.add_page_break()
    hp_hod = doc.add_paragraph()
    hrun_hod = hp_hod.add_run("HOD APPROVAL & REMARKS")
    hrun_hod.bold = True
    
    hod_table = doc.add_table(rows=2, cols=2)
    hod_table.style = 'Table Grid'
    hod_table.cell(0, 0).paragraphs[0].add_run("REMARKS:")
    hod_table.rows[1].height = Pt(50)
    hod_table.cell(1, 0).paragraphs[0].add_run("DATE:")
    hod_table.cell(1, 1).paragraphs[0].add_run("SIGNATURE:")

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# --- 4. STREAMLIT GRAPHICAL INTERFACE WORKFLOW ---
st.write("---")
st.info("INPUT THE LESSON TOPIC, THE SYLLABUS CODE, AND ANY DIGITAL TOOLS INTEGRATION CONTEXTS BELOW.")

# Inputs layout sequentially stacked underneath the API Input field
u_topic = st.text_input("LESSON TOPIC:")
u_syllabus = st.text_input("SYLLABUS CODE / SPECIFICATION:")
u_extra = st.text_area("EXTRA CONTEXT / DIGITAL TOOLS (OPTIONAL):")

if st.button("🚀 GENERATE PEDATI LESSON PLAN", type="primary"):
    if not user_api_key:
        st.error("❌ KEY CONFIGURATION ERROR! PLEASE INPUT YOUR PERSONAL GEMINI API KEY AT THE TOP FIRST.")
    elif not u_topic or not u_syllabus:
        st.error("❌ PLEASE FILL OUT BOTH LESSON TOPIC AND SYLLABUS CODE SPECIFICATIONS.")
    else:
        with st.spinner("AI IS STRUCTURING YOUR CUSTOM PEDATI LESSON PLAN..."):
            result = generate_pedati_plan(u_topic, u_syllabus, u_extra, user_api_key, selected_model_name)
            st.session_state['pedati_plan_out_en'] = result

if 'pedati_plan_out_en' in st.session_state:
    st.divider()
    st.subheader("👁️ AI DRAFT PREVIEW")
    st.text_area("GENERATED CONTENT CONTENT PREVIEW", st.session_state['pedati_plan_out_en'], height=450)
    
    doc_file = create_word_export(u_topic, u_syllabus, st.session_state['pedati_plan_out_en'])
    st.download_button(
        label="📥 DOWNLOAD TO OFFICIAL WORD VERSION (.DOCX)", 
        data=doc_file, 
        file_name=f"PEDATI_LP_{u_topic.upper().replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

st.markdown("---")
st.caption("PEDATI LESSON PLANNER 3.0 | DEVELOPER: HJH NURUL HAZIQAH HJ NORDIN | © 2026 BRUNEI'S EDUCATION INNOVATION")
