import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn
import os
import math
import tempfile
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Instant Collage Maker", 
    page_icon="🖼️", 
    layout="centered"
)

# --- CSS STYLING FOR USER FRIENDLY UI ---
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #2E86C1;
        text-align: center;
        font-weight: bold;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #566573;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stButton>button {
        color: white;
        background-color: #2E86C1;
        border-radius: 5px;
        height: 3em;
        width: 100%;
        font-size: 1.1rem;
        font-weight: bold;
    }
    .stDownloadButton>button {
        background-color: #28B463;
        color: white;
        border-radius: 5px;
        height: 3em;
        width: 100%;
        font-size: 1.1rem;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# --- HELPER FUNCTIONS (Word Logic) ---
def set_cell_margins(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin in ["top", "start", "bottom", "end"]:
        if margin in kwargs:
            node = OxmlElement(f'w:{margin}')
            node.set(qn('w:w'), str(kwargs[margin]))
            tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_vertical_align(cell, align="center"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), align)
    tcPr.append(vAlign)

def set_cell_border(cell):
    """Adds a simple black border to the cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('start', 'top', 'end', 'bottom'):
        element = OxmlElement(f'w:{edge}')
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), '12') # Border thickness
        element.set(qn('w:color'), '000000') # Black color
        tcBorders.append(element)
    tcPr.append(tcBorders)

# --- GENERATOR FUNCTIONS ---

def create_word_doc(images, title):
    doc = Document()
    
    # Page Setup
    section = doc.sections[0]
    page_w, page_h = 8.0, 11.0
    num = len(images)
    
    if num == 0: return None

    # Title
    if title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        r.font.size = Pt(24)
        r.bold = True
        p.space_after = Pt(6)
        h_offset = 0.6
    else:
        h_offset = 0
        
    avail_h = page_h - h_offset

    # Grid Logic
    if num == 1: c, r = 1, 1
    elif num <= 4: c, r = 2, 2
    elif num <= 9: c, r = 3, 3
    elif num <= 16: c, r = 4, 4
    else:
        c = math.ceil(math.sqrt(num))
        r = math.ceil(num / c)
        
    cell_w = (page_w / c)
    cell_h = (avail_h / r)

    # Build Table
    table = doc.add_table(rows=r, cols=c)
    table.autofit = False
    for col in table.columns: col.width = Inches(page_w / c)

    idx = 0
    for i in range(r):
        for j in range(c):
            if idx >= num: break
            
            # Add Image to Cell
            cell = table.rows[i].cells[j]
            set_cell_margins(cell, top=0, start=0, bottom=0, end=0)
            set_cell_vertical_align(cell, "center")
            set_cell_border(cell)
            
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Temp file for editable word objects
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                tmp.write(images[idx])
                tmp_name = tmp.name
            
            try:
                p.add_run().add_picture(tmp_name, width=Inches(cell_w), height=Inches(cell_h))
            except: pass
            
            os.remove(tmp_name)
            idx += 1

    # Save
    f = BytesIO()
    doc.save(f)
    f.seek(0)
    return f

def create_png(images, title):
    num = len(images)
    if num == 0: return None

    # Canvas Size (A4 @ 300 DPI)
    W, H = 2480, 3508
    
    # Title Space
    t_h = 250 if title else 0
    avail_h = H - t_h
    
    canvas = Image.new('RGB', (W, H), 'white')
    
    # Draw Title
    if title:
        draw = ImageDraw.Draw(canvas)
        try: font = ImageFont.truetype("arial.ttf", 80)
        except: font = ImageFont.load_default()
        
        bbox = draw.textbbox((0, 0), title, font=font)
        txt_w = bbox[2] - bbox[0]
        txt_h = bbox[3] - bbox[1]
        
        draw.text(((W-txt_w)/2, (t_h-txt_h)/2), title, font=font, fill="black")

    # Grid Logic
    if num == 1: c, r = 1, 1
    elif num <= 4: c, r = 2, 2
    elif num <= 9: c, r = 3, 3
    elif num <= 16: c, r = 4, 4
    else:
        c = math.ceil(math.sqrt(num))
        r = math.ceil(num / c)
        
    cw = W // c
    ch = avail_h // r
    
    idx = 0
    for i in range(r):
        for j in range(c):
            if idx >= num: break
            try:
                img = Image.open(BytesIO(images[idx]))
                img = img.resize((cw, ch), Image.Resampling.LANCZOS)
                
                # Draw Border
                d = ImageDraw.Draw(img)
                d.rectangle([0, 0, cw, ch], outline="black", width=5)
                
                # Paste
                canvas.paste(img, (j*cw, t_h + i*ch))
            except: pass
            idx += 1
            
    f = BytesIO()
    canvas.save(f, format="PNG")
    f.seek(0)
    return f

# --- MAIN APP UI ---

st.markdown('<p class="main-header">Instant Collage Maker</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Upload photos • Add Title • Download</p>', unsafe_allow_html=True)

with st.container():
    title_input = st.text_input("📝 Collage Title", placeholder="e.g., Summer Vacation 2024", label_visibility="collapsed")

st.divider()

uploaded_files = st.file_uploader(
    "📷 Upload your images (JPG, PNG)", 
    type=['png', 'jpg', 'jpeg'], 
    accept_multiple_files=True,
    label_visibility="visible"
)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} images loaded. Ready to generate!")
    
    raw_imgs = [f.read() for f in uploaded_files]
    
    col1, col2 = st.columns(2, gap="large")
    
    with col1:
        st.markdown("### 📄 Word Document")
        st.caption("Best for editing text or moving images later.")
        if st.button("Generate Word", use_container_width=True):
            with st.spinner("Creating Word file..."):
                res = create_word_doc(raw_imgs, title_input)
                if res:
                    st.download_button("⬇️ Download .docx", res, f"{title_input or 'Collage'}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    with col2:
        st.markdown("### 🖼️ PNG Image")
        st.caption("Best for printing or sharing on social media.")
        if st.button("Generate PNG", use_container_width=True):
            with st.spinner("Rendering High-Res Image..."):
                res = create_png(raw_imgs, title_input)
                if res:
                    st.download_button("⬇️ Download .png", res, f"{title_input or 'Collage'}.png", mime="image/png")

else:
    st.info("👈 Please upload images to get started.")

# Footer
st.markdown("---")
st.markdown("<center><small>Made with ❤️ using Python & Streamlit</small></center>", unsafe_allow_html=True)