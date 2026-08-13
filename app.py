import streamlit as st
import fitz
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io

st.set_page_config(
    page_title="PDF to PPT Generator",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF to Editable PPT Generator")
st.write("Convert PDF text into editable PowerPoint slides.")

st.divider()

pdf_file = st.file_uploader(
    "Upload your PDF",
    type=["pdf"]
)

if pdf_file:

    st.success(f"PDF selected: {pdf_file.name}")

    if st.button("🚀 Generate Editable PowerPoint", type="primary"):

        with st.spinner("Reading PDF and creating editable PPT..."):

            pdf_bytes = pdf_file.read()
            pdf = fitz.open(
                stream=pdf_bytes,
                filetype="pdf"
            )

            prs = Presentation()

            # 16:9 widescreen
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)

            for page in pdf:

                slide = prs.slides.add_slide(
                    prs.slide_layouts[6]
                )

                page_width = page.rect.width
                page_height = page.rect.height

                # Extract text blocks
                blocks = page.get_text("blocks")

                for block in blocks:

                    x0, y0, x1, y1, text = block[:5]

                    text = text.strip()

                    if not text:
                        continue

                    # Convert PDF coordinates to PPT coordinates
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

            output = io.BytesIO()

            prs.save(output)

            output.seek(0)

        st.success(
            "✅ Editable PowerPoint created successfully!"
        )

        st.download_button(
            label="⬇️ Download Editable PowerPoint",
            data=output,
            file_name="editable_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
