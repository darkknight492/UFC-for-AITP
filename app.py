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
        st.error(f"System Error: Could not save file to disk. Details: {e}")
        return None

def convert_file(file_path):
    """
    Orchestrates the conversion.
    Returns: (Boolean Success, String Result/Error)
    """
    try:
        # The primary engine call
        result = md_engine.convert(file_path)
        
        # Check if result is empty or None
        if result is None or result.text_content is None:
             return False, "The converter returned no text. The file might be empty or image-only."
             
        return True, result.text_content
        
    except Exception as e:
        # We capture the exact error message here to show the user
        return False, str(e)

# --- Main Interface ---

st.title("📄 Universal File-to-Markdown")
st.markdown(
    """
    Convert **Word, Excel, PowerPoint, PDF, and HTML** into clean text.
    """
)

# [2] The Interface: Upload Area
uploaded_files = st.file_uploader(
    "Drag and drop files here", 
    accept_multiple_files=True,
    # We allow all likely extensions
    type=['docx', 'xlsx', 'pptx', 'pdf', 'html', 'htm', 'csv', 'txt']
)

if uploaded_files:
    st.divider()
    st.subheader("Processing Queue")

    for uploaded_file in uploaded_files:
        # Create a collapsible status container
        with st.status(f"Processing **{uploaded_file.name}**...", expanded=False) as status:
            
            # 1. Save to temp file
            temp_path = save_uploaded_file(uploaded_file)
            
            if temp_path:
                # 2. Perform Conversion with Error Catching
                success, output_text = convert_file(temp_path)
                
                # 3. Clean up temp file immediately
                try:
                    os.remove(temp_path)
                except:
                    pass # Non-critical error if delete fails

                # [3] Resilience: Check success
                if not success:
                    status.update(label=f"❌ Failed: {uploaded_file.name}", state="error")
                    st.error(f"⚠️ **Error Details:** {output_text}")
                    st.caption("Tip: If the error says 'ModuleNotFoundError', check your requirements.txt.")
                else:
                    status.update(label=f"✅ Ready: {uploaded_file.name}", state="complete")
                    
                    # Generate Output Filename
                    base_name = os.path.splitext(uploaded_file.name)[0]
                    output_name = f"{base_name}_converted"
                    
                    # [2] The Interface: Instant Preview & Download
                    with st.expander(f"👁️ Preview Content: {base_name}", expanded=True):
                        
                        # Tabs for Markdown vs Raw Text
                        tab1, tab2 = st.tabs(["Rendered View", "Raw Text"])
                        with tab1:
                            st.markdown(output_text)
                        with tab2:
                            st.text_area("Copy content", output_text, height=200)

                        st.divider()
                        
                        # Download Buttons
                        col1, col2 = st.columns(2)
                        with col1:
                            st.download_button(
                                label="📥 Download as Markdown (.md)",
                                data=output_text,
                                file_name=f"{output_name}.md",
                                mime="text/markdown"
                            )
                        with col2:
                            st.download_button(
                                label="📥 Download as Text (.txt)",
                                data=output_text,
                                file_name=f"{output_name}.txt",
                                mime="text/plain"
                            )

# Footer
st.divider()
st.caption("Powered by Microsoft MarkItDown")
