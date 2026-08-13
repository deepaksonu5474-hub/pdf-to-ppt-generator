import streamlit as st
import fitz
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from PIL import Image
import io
import os

st.set_page_config(
    page_title="PDF to PPT Generator",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF to Editable PPT Generator")
st.write("Convert PDF, DOC, DOCX and Images into PowerPoint.")

st.divider()

uploaded_file = st.file_uploader(
    "Upload your file",
    type=[
        "pdf",
        "doc",
        "docx",
        "png",
        "jpg",
        "jpeg",
        "webp"
    ]
)

if uploaded_file:

    file_name = uploaded_file.name
    file_ext = os.path.splitext(file_name)[1].lower()

    st.success(f"File selected: {file_name}")

    if st.button(
        "🚀 Generate PowerPoint",
        type="primary"
    ):

        with st.spinner(
            "Reading file and creating PowerPoint..."
        ):

            file_bytes = uploaded_file.read()

            prs = Presentation()

            # 16:9 widescreen
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)

            # ------------------------------------------------
            # PDF
            # ------------------------------------------------

            if file_ext == ".pdf":

                pdf = fitz.open(
                    stream=file_bytes,
                    filetype="pdf"
                )

                for page in pdf:

                    slide = prs.slides.add_slide(
                        prs.slide_layouts[6]
                    )

                    page_width = page.rect.width
                    page_height = page.rect.height

                    blocks = page.get_text("blocks")

                    for block in blocks:

                        x0, y0, x1, y1, text = block[:5]

                        text = text.strip()

                        if not text:
                            continue

                        left = Inches(
                            (x0 / page_width) * 13.333
                        )

                        top = Inches(
                            (y0 / page_height) * 7.5
                        )

                        width = Inches(
                            ((x1 - x0) / page_width) * 13.333
                        )

                        height = Inches(
                            ((y1 - y0) / page_height) * 7.5
                        )

                        textbox = slide.shapes.add_textbox(
                            left,
                            top,
                            width,
                            height
                        )

                        text_frame = textbox.text_frame
                        text_frame.clear()

                        p = text_frame.paragraphs[0]

                        run = p.add_run()
                        run.text = text

                        run.font.size = Pt(16)

                        p.alignment = PP_ALIGN.LEFT

            # ------------------------------------------------
            # IMAGE
            # ------------------------------------------------

            elif file_ext in [
                ".png",
                ".jpg",
                ".jpeg",
                ".webp"
            ]:

                slide = prs.slides.add_slide(
                    prs.slide_layouts[6]
                )

                image = Image.open(
                    io.BytesIO(file_bytes)
                )

                image_buffer = io.BytesIO()

                image.convert("RGB").save(
                    image_buffer,
                    format="PNG"
                )

                image_buffer.seek(0)

                slide.shapes.add_picture(
                    image_buffer,
                    0,
                    0,
                    width=prs.slide_width,
                    height=prs.slide_height
                )

            # ------------------------------------------------
            # DOC / DOCX
            # ------------------------------------------------

            elif file_ext in [
                ".doc",
                ".docx"
            ]:

                slide = prs.slides.add_slide(
                    prs.slide_layouts[6]
                )

                textbox = slide.shapes.add_textbox(
                    Inches(0.5),
                    Inches(0.5),
                    Inches(12.3),
                    Inches(6.5)
                )

                text_frame = textbox.text_frame

                text_frame.text = (
                    "DOC/DOCX file uploaded successfully.\n\n"
                    "DOC/DOCX → Editable PPT conversion "
                    "will be added in the next version."
                )

                for paragraph in text_frame.paragraphs:

                    for run in paragraph.runs:

                        run.font.size = Pt(20)

            # ------------------------------------------------
            # SAVE PPT
            # ------------------------------------------------

            output = io.BytesIO()

            prs.save(output)

            output.seek(0)

        st.success(
            "✅ PowerPoint created successfully!"
        )

        st.download_button(
            label="⬇️ Download PowerPoint",
            data=output,
            file_name="converted_presentation.pptx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "presentationml.presentation"
            )
        )
