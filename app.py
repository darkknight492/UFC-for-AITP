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
@st.cache_resource
def get_converter():
    return MarkItDown()

md_engine = get_converter()

# --- Helper Functions ---

def format_size(size_in_bytes):
    """Converts bytes to a readable string (e.g., 1048576 -> '1.00 MB')."""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_in_bytes < 1024.0:
            return f"{size_in_bytes:.2f} {unit}"
        size_in_bytes /= 1024.0
    return f"{size_in_bytes:.2f} TB"

def save_uploaded_file(uploaded_file):
    """Saves upload to temp file to preserve extension."""
    try:
        suffix = os.path.splitext(uploaded_file.name)[1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            return tmp_file.name
    except Exception as e:
        st.error(f"System Error: Could not save file. Details: {e}")
        return None

def convert_file(file_path):
    """Orchestrates the conversion. Returns (Success, Result/Error)."""
    try:
        result = md_engine.convert(file_path)
        if result is None or result.text_content is None:
             return False, "The converter returned no text. File might be empty/image-only."
        return True, result.text_content
    except Exception as e:
        return False, str(e)

# --- Main Interface ---

st.title("📄 Universal File-to-Markdown")
st.markdown("Convert **Word, Excel, PowerPoint, PDF, and HTML** into clean text.")

uploaded_files = st.file_uploader(
    "Drag and drop files here", 
    accept_multiple_files=True,
    type=['docx', 'xlsx', 'pptx', 'pdf', 'html', 'htm', 'csv', 'txt']
)

if uploaded_files:
    st.divider()
    st.subheader("Processing Queue")

    for uploaded_file in uploaded_files:
        with st.status(f"Processing **{uploaded_file.name}**...", expanded=False) as status:
            
            temp_path = save_uploaded_file(uploaded_file)
            
            if temp_path:
                success, output_text = convert_file(temp_path)
                
                try:
                    os.remove(temp_path)
                except:
                    pass

                if not success:
                    status.update(label=f"❌ Failed: {uploaded_file.name}", state="error")
                    st.error(f"⚠️ **Error Details:** {output_text}")
                else:
                    status.update(label=f"✅ Ready: {uploaded_file.name}", state="complete")
                    
                    # --- Logic: File Size Calculations ---
                    original_bytes = uploaded_file.size
                    # We estimate text size by encoding strictly to UTF-8
                    converted_bytes = len(output_text.encode('utf-8'))
                    
                    original_fmt = format_size(original_bytes)
                    converted_fmt = format_size(converted_bytes)
                    
                    # Calculate percentage reduction
                    if original_bytes > 0:
                        reduction_pct = ((original_bytes - converted_bytes) / original_bytes) * 100
                    else:
                        reduction_pct = 0
                    
                    base_name = os.path.splitext(uploaded_file.name)[0]
                    output_name = f"{base_name}_converted"
                    
                    # --- Interface: Tabs ---
                    with st.expander(f"👁️ Preview Content: {base_name}", expanded=True):
                        
                        # Tabs for Rendered View, Raw Text, and Stats
                        tab1, tab2, tab3 = st.tabs(["Rendered View", "Raw Text", "📊 File Size Comparison"])
                        
                        with tab1:
                            st.markdown(output_text)
                        
                        with tab2:
                            st.text_area("Copy content", output_text, height=200)
                            
                        with tab3:
                            # Display Size Comparison Metrics
                            col_a, col_b = st.columns(2)
                            col_a.metric("Original Size", original_fmt)
                            col_b.metric("Text Size", converted_fmt, delta=f"-{reduction_pct:.1f}% reduction")
                            
                            st.divider()
                            st.markdown(f"**Efficiency:** Text version is **{reduction_pct:.1f}% smaller** than the original.")
                            
                            # Clean Table
                            st.table({
                                "Metric": ["Original File Size", "Converted Text Size"],
                                "Value": [original_fmt, converted_fmt]
                            })

                        st.divider()
                        
                        # Download Buttons
                        col1, col2 = st.columns(2)
                        with col1:
                            st.download_button(
                                label="📥 Download .md",
                                data=output_text,
                                file_name=f"{output_name}.md",
                                mime="text/markdown"
                            )
                        with col2:
                            st.download_button(
                                label="📥 Download .txt",
                                data=output_text,
                                file_name=f"{output_name}.txt",
                                mime="text/plain"
                            )

st.divider()
st.caption("Powered by Microsoft MarkItDown")
