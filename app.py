import streamlit as st
import os
import tempfile
from markitdown import MarkItDown

# --- Configuration & Setup ---
st.set_page_config(
    page_title="Universal File-to-Text Converter",
    page_icon="📄",
    layout="centered"
)

# Initialize the conversion engine
# We cache this to avoid reloading heavy libraries on every interaction
@st.cache_resource
def get_converter():
    return MarkItDown()

md_engine = get_converter()

# --- Helper Functions ---

def save_uploaded_file(uploaded_file):
    """
    Saves the uploaded file to a temporary location while preserving 
    its extension so MarkItDown can identify the file type.
    """
    try:
        # Create a temp file with the same suffix (extension) as the upload
        suffix = os.path.splitext(uploaded_file.name)[1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            return tmp_file.name
    except Exception as e:
        st.error(f"Error handling file system: {e}")
        return None

def convert_file(file_path):
    """
    Orchestrates the conversion using MarkItDown.
    Includes basic error handling logic.
    """
    try:
        # The primary engine call
        result = md_engine.convert(file_path)
        return result.text_content
    except Exception as e:
        # Return None to signal failure to the UI
        return None

# --- Main Interface ---

st.title("📄 Universal File-to-Markdown")
st.markdown(
    """
    Convert your documents (**Word, Excel, PowerPoint, PDF, HTML**) into clean, portable Markdown text.
    """
)

# [2] The Interface: Upload Area
uploaded_files = st.file_uploader(
    "Drag and drop files here", 
    accept_multiple_files=True,
    type=['docx', 'xlsx', 'pptx', 'pdf', 'html', 'htm']
)

if uploaded_files:
    st.divider()
    st.subheader("Processing Queue")

    for uploaded_file in uploaded_files:
        with st.status(f"Converting **{uploaded_file.name}**...", expanded=False) as status:
            
            # 1. Save to temp file
            temp_path = save_uploaded_file(uploaded_file)
            
            if temp_path:
                # 2. Perform Conversion
                converted_text = convert_file(temp_path)
                
                # 3. Clean up temp file
                os.remove(temp_path)

                # [3] Resilience: Check if conversion was successful
                if converted_text is None:
                    status.update(label=f"❌ Failed: {uploaded_file.name}", state="error")
                    st.error(f"⚠️ Could not read {uploaded_file.name}. Please check the format.")
                else:
                    status.update(label=f"✅ Ready: {uploaded_file.name}", state="complete")
                    
                    # Generate Output Filename
                    base_name = os.path.splitext(uploaded_file.name)[0]
                    output_name = f"{base_name}_converted"
                    
                    # [2] The Interface: Instant Preview
                    with st.expander(f"👁️ Preview: {base_name}"):
                        st.markdown(converted_text) # Rendered Markdown
                        st.divider()
                        st.text_area("Raw Text Source", converted_text, height=200)

                    # [2] The Interface: Download Options
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.download_button(
                            label="📥 Download .md",
                            data=converted_text,
                            file_name=f"{output_name}.md",
                            mime="text/markdown"
                        )
                    
                    with col2:
                        st.download_button(
                            label="📥 Download .txt",
                            data=converted_text,
                            file_name=f"{output_name}.txt",
                            mime="text/plain"
                        )
            else:
                 st.error("System error: Could not save temporary file.")

# Footer / Instructions
st.divider()
st.caption("Powered by Microsoft MarkItDown & Streamlit")
