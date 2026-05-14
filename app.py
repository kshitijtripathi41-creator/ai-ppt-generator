import streamlit as st
import pandas as pd
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from markitdown import MarkItDown
import os
import io
import json
import re
import tempfile

# --- Configuration & Security ---
# Ensure GEMINI_API_KEY is set in your Codespaces environment variables
api_key = os.getenv("enterapikey")
if not api_key:
    st.error("🚨 GEMINI_API_KEY environment variable not found. Please set it in your Codespaces environment.")
    st.stop()

genai.configure(api_key=api_key)

# --- Helper Functions ---

def parse_excel(uploaded_file):
    """Reads all sheets from an Excel file and converts to text."""
    df_dict = pd.read_excel(uploaded_file, sheet_name=None)
    text_content = ""
    for sheet_name, df in df_dict.items():
        text_content += f"\\n--- Sheet: {sheet_name} ---\\n"
        text_content += df.to_string(index=False)
    return text_content

def parse_word(uploaded_file):
    """Uses MarkItDown to extract text from a Word document via a temp file."""
    md = MarkItDown()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    try:
        result = md.convert(tmp_path)
        text_content = result.text_content
    finally:
        os.remove(tmp_path) # Clean up temp file
    return text_content

def generate_presentation_data(data_text):
    """Calls Gemini to analyze data and return structured JSON for the slides."""
    model = genai.GenerativeModel("gemini-1.5-pro")
    
    prompt = f"""
    Analyze the following raw business data. Identify the top 5 strategic insights. 
    Structure a 7-slide PowerPoint presentation strictly including:
    1. Title Slide (Title and a brief subtitle)
    2. Executive Summary (High-level overview)
    3 to 6. Four distinct Insight Slides (Each with a title and 3-4 bullet points)
    7. Conclusion (Final thoughts and next steps)

    You MUST return ONLY valid JSON in the exact format below, with no markdown wrappers or additional text:
    [
        {{"title": "Presentation Title", "content": ["Subtitle"]}},
        {{"title": "Executive Summary", "content": ["Point 1", "Point 2"]}},
        {{"title": "Insight 1 Title", "content": ["Bullet 1", "Bullet 2"]}},
        ...
    ]

    Data:
    {data_text}
    """
    
    response = model.generate_content(prompt)
    
    # Clean up the response in case Gemini includes markdown code blocks
    cleaned_text = response.text.replace("```json", "").replace("```", "").strip()
    return json.loads(cleaned_text)

def create_pptx(slides_data):
    """Generates a PowerPoint presentation using python-pptx with Modern Corporate styling."""
    prs = Presentation()
    
    # Modern Corporate Colors
    DARK_BLUE = RGBColor(0, 51, 102)

    for i, slide_dict in enumerate(slides_data):
        title_text = slide_dict.get("title", "Slide Title")
        content_list = slide_dict.get("content", [])

        if i == 0:
            # Title Slide
            slide_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(slide_layout)
            title = slide.shapes.title
            subtitle = slide.placeholders[1]
            
            title.text = title_text
            title.text_frame.paragraphs[0].font.name = 'Calibri'
            title.text_frame.paragraphs[0].font.color.rgb = DARK_BLUE
            title.text_frame.paragraphs[0].font.bold = True
            
            if content_list:
                subtitle.text = content_list[0]
                subtitle.text_frame.paragraphs[0].font.name = 'Calibri'
        else:
            # Content Slide
            slide_layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(slide_layout)
            title = slide.shapes.title
            body_shape = slide.placeholders[1]
            
            title.text = title_text
            title.text_frame.paragraphs[0].font.name = 'Calibri'
            title.text_frame.paragraphs[0].font.color.rgb = DARK_BLUE
            title.text_frame.paragraphs[0].font.bold = True
            
            tf = body_shape.text_frame
            tf.clear() # Clear default paragraph
            
            for bullet in content_list:
                p = tf.add_paragraph()
                p.text = bullet
                p.font.name = 'Calibri'
                p.font.size = Pt(18)
                p.space_after = Pt(14) # Plenty of whitespace

    # Save to an in-memory buffer
    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- Streamlit UI ---

st.set_page_config(page_title="AI Data to PPTX Builder", page_icon="📊", layout="centered")

st.title("📊 Business Data to PowerPoint AI")
st.markdown("Upload your raw `.xlsx` or `.docx` files. Gemini will analyze the data, extract the top strategic insights, and build a modern corporate presentation.")

uploaded_file = st.file_uploader("Upload Data File", type=["xlsx", "docx"])

if uploaded_file is not None:
    if st.button("Generate Presentation"):
        # Initialize progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()

        try:
            # Step 1: Parsing Data
            status_text.text("Parsing uploaded file...")
            if uploaded_file.name.endswith(".xlsx"):
                raw_text = parse_excel(uploaded_file)
            else:
                raw_text = parse_word(uploaded_file)
            progress_bar.progress(30)

            # Step 2: AI Analysis
            status_text.text("Analyzing data with Gemini and generating insights...")
            slides_data = generate_presentation_data(raw_text)
            progress_bar.progress(70)

            # Step 3: PPTX Generation
            status_text.text("Applying Modern Corporate styling and rendering slides...")
            ppt_file = create_pptx(slides_data)
            progress_bar.progress(100)
            status_text.text("✅ Presentation generated successfully!")

            # Step 4: Download Button
            st.download_button(
                label="📥 Download Presentation",
                data=ppt_file,
                file_name="Strategic_Insights_Deck.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        except Exception as e:
            st.error(f"An error occurred: {e}")
            progress_bar.empty()
