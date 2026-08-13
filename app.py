import streamlit as st
import fitz
from pptx import Presentation
from pptx.util import Inches
import io

st.set_page_config(
    page_title="PDF to PPT Generator",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF to Editable PPT Generator")
st.write("Convert your PDF into a PowerPoint presentation.")

st.divider()

pdf_file = st.file_uploader(
    "Upload your PDF",
    type=["pdf"]
)

if pdf_file:

    st.success(f"PDF selected: {pdf_file.name}")

    if st.button("🚀 Generate PowerPoint", type="primary"):

        with st.spinner("Creating PowerPoint..."):

            pdf_bytes = pdf_file.read()
            pdf = fitz.open(stream=pdf_bytes, filetype="pdf")

            prs = Presentation()

            # 16:9 widescreen
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)

            for page in pdf:

                slide = prs.slides.add_slide(
                    prs.slide_layouts[6]
                )

                # Render PDF page as image
                pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))

                image_bytes = pix.tobytes("png")

                slide.shapes.add_picture(
                    io.BytesIO(image_bytes),
                    0,
                    0,
                    width=prs.slide_width,
                    height=prs.slide_height
                )

            output = io.BytesIO()
            prs.save(output)
            output.seek(0)

        st.success("✅ PowerPoint created successfully!")

        st.download_button(
            label="⬇️ Download PowerPoint",
            data=output,
            file_name="converted_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
