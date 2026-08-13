import streamlit as st

st.set_page_config(
    page_title="PDF to PPT Generator",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF to Editable PPT Generator")
st.write("Convert your PDF into an editable PowerPoint presentation.")

st.divider()

pdf_file = st.file_uploader(
    "Upload your PDF",
    type=["pdf"]
)

if pdf_file:
    st.success(f"PDF selected: {pdf_file.name}")

    st.info(
        "PDF processing engine will be connected in the next step."
    )

    if st.button("🚀 Generate PowerPoint", type="primary"):
        st.warning("Conversion engine is not connected yet.")
