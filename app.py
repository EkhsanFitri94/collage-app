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
    page_title="Collage App", 
    page_icon="🖼️", 
    layout="centered"
)

# --- MOBILE-FRIENDLY CSS ---
st.markdown("""
<style>
    /* Make text responsive */
    h1 { font-size: 1.8rem; text-align: center; color: #2E86C1; }
    
    /* Make buttons large and easy to tap on mobile */
    div.stButton > button {
        width: 100%;
        height: 3.5em;
        font-size: 1.1rem;
        font-weight: bold;
        border-radius: 10px;
        margin-top: 10px;
        margin-bottom: 10px;
    }
    
    /* Color specific buttons */
    .generate-word { background-color: #2E86C1; color: white; border: none; }
    .generate-png { background-color: #E67E22; color: white; border: none; }
    .camera-btn { background-color: #28B463; color: white; border: none; }
    
    /* Center the file uploader */
    [data-testid="stFileUpload"] {
        text-align: center;
    }
    
    /* Reduce padding on mobile */
    .block-container { padding-top: 1rem; padding-bottom: 1rem; }
</style>
""", unsafe_allow_html=True)

# --- HELPER FUNCTIONS ---

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
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('start', 'top', 'end', 'bottom'):
        element = OxmlElement(f'w:{edge}')
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), '12') 
        element.set(qn('w:color'), '000000')
        tcBorders.append(element)
    tcPr.append(tcBorders)

# --- GENERATORS ---

def create_word_doc(images, title):
    doc = Document()
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

    # Grid
    if num == 1: c, r = 1, 1
    elif num <= 4: c, r = 2, 2
    elif num <= 9: c, r = 3, 3
    elif num <= 16: c, r = 4, 4
    else:
        c = math.ceil(math.sqrt(num))
        r = math.ceil(num / c)
        
    cell_w = (page_w / c)
    cell_h = (avail_h / r)

    table = doc.add_table(rows=r, cols=c)
    table.autofit = False
    for col in table.columns: col.width = Inches(page_w / c)

    idx = 0
    for i in range(r):
        for j in range(c):
            if idx >= num: break
            
            cell = table.rows[i].cells[j]
            set_cell_margins(cell, top=0, start=0, bottom=0, end=0)
            set_cell_vertical_align(cell, "center")
            set_cell_border(cell)
            
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                tmp.write(images[idx])
                tmp_name = tmp.name
            
            try:
                p.add_run().add_picture(tmp_name, width=Inches(cell_w), height=Inches(cell_h))
            except: pass
            
            os.remove(tmp_name)
            idx += 1

    f = BytesIO()
    doc.save(f)
    f.seek(0)
    return f

def create_png(images, title):
    num = len(images)
    if num == 0: return None

    W, H = 2480, 3508
    t_h = 250 if title else 0
    avail_h = H - t_h
    
    canvas = Image.new('RGB', (W, H), 'white')
    
    if title:
        draw = ImageDraw.Draw(canvas)
        try: font = ImageFont.truetype("arial.ttf", 80)
        except: font = ImageFont.load_default()
        
        bbox = draw.textbbox((0, 0), title, font=font)
        txt_w = bbox[2] - bbox[0]
        txt_h = bbox[3] - bbox[1]
        
        draw.text(((W-txt_w)/2, (t_h-txt_h)/2), title, font=font, fill="black")

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
                d = ImageDraw.Draw(img)
                d.rectangle([0, 0, cw, ch], outline="black", width=5)
                canvas.paste(img, (j*cw, t_h + i*ch))
            except: pass
            idx += 1
            
    f = BytesIO()
    canvas.save(f, format="PNG")
    f.seek(0)
    return f

# --- UI ---

# 1. Header
st.markdown("<h1>Mobile Collage Maker</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center; color:#666;'>Tap photos or use camera</p>", unsafe_allow_html=True)

# 2. Input
title_input = st.text_input("📝 Title", placeholder="e.g. My Trip", label_visibility="collapsed")

# 3. Upload Options
# We use a column layout to put File Upload and Camera side-by-side on desktop, 
# but they will stack automatically on mobile.
up_col, cam_col = st.columns(2)

with up_col:
    uploaded_files = st.file_uploader(
        "📂 Gallery", 
        type=['png', 'jpg', 'jpeg'], 
        accept_multiple_files=True,
        label_visibility="collapsed"
    )

with cam_col:
    # This specific button opens the phone camera directly
    camera_files = st.camera_input("📸 Camera", label_visibility="collapsed")

# Combine images from both sources
all_images_raw = []
if uploaded_files:
    all_images_raw.extend([f.read() for f in uploaded_files])
if camera_files:
    # camera_input returns a single file, handle it if multiple photos aren't taken
    # Note: Standard st.camera_input takes one photo at a time.
    all_images_raw.append(camera_files.read())

if all_images_raw:
    st.success(f"✅ {len(all_images_raw)} photo(s) selected")
    
    # Use full width for buttons on mobile for easier tapping
    st.markdown("---")
    
    # Generate Buttons
    st.markdown("### Choose Format")
    
    res_doc = create_word_doc(all_images_raw, title_input)
    res_png = create_png(all_images_raw, title_input)
    
    # Button Column 1: Word
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if res_doc:
            st.download_button(
                "📄 Word Doc", 
                res_doc, 
                f"{title_input or 'Collage'}.docx", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_word"
            )
            
    with col_btn2:
        if res_png:
            st.download_button(
                "🖼️ PNG Image", 
                res_png, 
                f"{title_input or 'Collage'}.png", 
                mime="image/png",
                key="download_png"
            )

else:
    st.info("👇 Upload from gallery or take a photo below")

st.markdown("<center><small>Mobile Optimized</small></center>", unsafe_allow_html=True)
