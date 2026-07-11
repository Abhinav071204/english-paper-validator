import streamlit as st
import os
import tempfile
from validator import DocumentValidator, auto_correct_document

# Set page configuration
st.set_page_config(
    page_title="ELA Paper Validator",
    page_icon="✍️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern visual design and inline highlights
st.markdown("""
<style>
    .highlight-placeholder {
        background-color: rgba(220, 38, 38, 0.08);
        border-bottom: 2px dotted #dc2626;
        cursor: help;
        font-weight: 500;
    }
    .highlight-font {
        background-color: rgba(217, 119, 6, 0.08);
        border-bottom: 2px dotted #d97706;
        cursor: help;
        font-weight: 500;
    }
    .highlight-wording {
        background-color: rgba(2, 132, 199, 0.08);
        border-bottom: 2px dotted #0284c7;
        cursor: help;
        font-weight: 500;
    }
    .highlight-grammar {
        background-color: rgba(124, 58, 237, 0.08);
        border-bottom: 2px dotted #7c3aed;
        cursor: help;
        font-weight: 500;
    }
    .highlight-block-error {
        border-left: 4px solid #dc2626;
        padding-left: 0.75rem;
        background-color: rgba(220, 38, 38, 0.02);
    }
    .highlight-block-warning {
        border-left: 4px solid #d97706;
        padding-left: 0.75rem;
        background-color: rgba(217, 119, 6, 0.02);
    }
    .doc-preview-canvas {
        background-color: #ffffff;
        color: #0f172a;
        padding: 3rem;
        border-radius: 8px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
        border: 1px solid #e2e8f0;
        font-family: Arial, sans-serif;
    }
    .doc-preview-canvas p {
        font-size: 11pt;
        line-height: 1.5;
        margin-bottom: 1.2rem;
    }
    .doc-preview-canvas p.heading-1 {
        font-size: 20pt;
        font-weight: bold;
        color: #000000;
        margin-top: 1.5rem;
        margin-bottom: 0.75rem;
    }
    .doc-preview-canvas p.heading-2 {
        font-size: 16pt;
        font-weight: bold;
        color: #000000;
        margin-top: 1.2rem;
        margin-bottom: 0.6rem;
    }
    .doc-preview-canvas p.heading-3 {
        font-size: 12pt;
        font-weight: bold;
        color: #000000;
        margin-top: 1rem;
        margin-bottom: 0.5rem;
    }
    .doc-preview-canvas table {
        border-collapse: collapse;
        width: 100%;
        margin: 1.5rem 0;
        font-size: 10pt;
    }
    .doc-preview-canvas td {
        border: 1px solid #cbd5e1;
        padding: 0.6rem 0.8rem;
    }
</style>
""", unsafe_allow_html=True)

# Helper function to escape HTML
def escape_html(text):
    if not text:
        return ''
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;').replace("'", '&#039;')

# Helper function to inject highlights in text
def apply_text_highlights(text, findings):
    if not text.strip():
        return ''
    escaped = escape_html(text)
    
    # Sort findings by length of target text descending
    run_findings = [f for f in findings if f.get("text")]
    run_findings.sort(key=lambda x: len(x["text"]), reverse=True)
    
    replaced = set()
    for f in run_findings:
        target = escape_html(f["text"])
        if not target or target in replaced:
            continue
            
        highlight_span = f'<span class="highlight-{f["category"]}" title="{escape_html(f["message"])}">{target}</span>'
        
        if target in escaped:
            escaped = escaped.replace(target, highlight_span)
            replaced.add(target)
            
    return escaped.replace('\n', '<br>')

st.title("✍️ ELA Document Validator & Auto-Corrector")
st.markdown("Upload your ELA test paper (`.docx`) to validate layout standards, check readability levels, and auto-correct formatting errors instantly.")

# Sidebar uploads
st.sidebar.header("Upload Document")
uploaded_file = st.sidebar.file_uploader("Choose a Word document", type=["docx"])

if uploaded_file is not None:
    # Save uploaded file to temp path
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        temp_path = tmp_file.name

    try:
        # Run validation
        validator = DocumentValidator(temp_path)
        result = validator.run_validation()
        
        # Display Compliance Score
        score = result["stats"]["overall_compliance_score"]
        
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("Compliance Score", f"{score}/100")
        with col2:
            st.metric("Font Mismatches", result["stats"]["font_violations_count"])
        with col3:
            st.metric("Placeholders Left", result["stats"]["placeholders_count"])
        with col4:
            st.metric("Wording Issues", result["stats"]["wording_violations_count"])
        with col5:
            st.metric("Grammar Issues", result["stats"]["grammar_violations_count"])
            
        # Download corrected file section
        st.markdown("### 🛠️ Document Actions")
        
        # Run auto-corrections to a temp path
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_corrected:
            corrected_path = tmp_corrected.name
            
        auto_correct_document(temp_path, corrected_path)
        
        with open(corrected_path, "rb") as f:
            corrected_bytes = f.read()
            
        corrected_name = uploaded_file.name.replace(".docx", "_Corrected.docx")
        st.download_button(
            label="⬇️ Download Auto-Corrected Version",
            data=corrected_bytes,
            file_name=corrected_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        # Sidebar readability scores
        st.sidebar.markdown("---")
        st.sidebar.subheader("Passage Readability")
        for passage, data in result["passages_readability"].items():
            st.sidebar.markdown(f"**{passage}**")
            st.sidebar.markdown(f"- Grade Level: **{data['grade']}**")
            st.sidebar.markdown(f"- Word Count: **{data['words']}**")
            st.sidebar.markdown(f"- Flesch Ease: **{data['ease']}**")
            
        # Main preview canvas tabs
        tab1, tab2 = st.tabs(["Interactive Preview", f"Detailed Findings List ({len(result['findings'])})"])
        
        with tab1:
            st.markdown("#### Document Layout Preview")
            st.caption("Hover over highlighted text to see the quality issues.")
            
            html_parts = ['<div class="doc-preview-canvas">']
            for el in result["html_preview"]:
                if el["type"] == "paragraph":
                    style_class = ""
                    style = el["style"].lower()
                    if style.startswith("heading 1"):
                        style_class = ' class="heading-1"'
                    elif style.startswith("heading 2"):
                        style_class = ' class="heading-2"'
                    elif style.startswith("heading 3"):
                        style_class = ' class="heading-3"'
                        
                    p_text = apply_text_highlights(el["text"], el["findings"])
                    
                    block_findings = [f for f in el["findings"] if not f.get("text")]
                    if block_findings:
                        has_error = any(f["level"] == "error" for f in block_findings)
                        block_class = "highlight-block-error" if has_error else "highlight-block-warning"
                        tooltip = "\\n".join(f["message"] for f in block_findings)
                        html_parts.append(f'<p{style_class} class="{block_class}" title="{tooltip}">{p_text}</p>')
                    else:
                        html_parts.append(f'<p{style_class}>{p_text}</p>')
                        
                elif el["type"] == "table":
                    table_style = ""
                    block_findings = [f for f in el["findings"] if not f.get("text")]
                    if block_findings:
                        has_error = any(f["level"] == "error" for f in block_findings)
                        block_class = "highlight-block-error" if has_error else "highlight-block-warning"
                        tooltip = "\\n".join(f["message"] for f in block_findings)
                        table_style = f' class="{block_class}" title="{tooltip}"'
                        
                    html_parts.append(f'<table{table_style}><tbody>')
                    for row in el["cells"]:
                        html_parts.append("<tr>")
                        for cell in row:
                            cell_text = apply_text_highlights(cell["text"], cell["findings"])
                            html_parts.append(f'<td>{cell_text}</td>')
                        html_parts.append("</tr>")
                    html_parts.append("</tbody></table>")
                    
            html_parts.append('</div>')
            st.markdown("".join(html_parts), unsafe_allow_html=True)
            
        with tab2:
            st.markdown("#### QA Findings Summary")
            if not result["findings"]:
                st.success("Congratulations! No compliance errors found in this document.")
            else:
                for f in result["findings"]:
                    level_badge = "🔴 ERROR" if f["level"] == "error" else "🟡 WARNING" if f["level"] == "warning" else "🔵 INFO"
                    elem_type = f["element_type"].upper() if f.get("element_type") else "DOCUMENT"
                    index_val = f"#{f['index']}" if f.get("index") is not None else ""
                    
                    st.markdown(f"**{level_badge}** - {elem_type} {index_val} ({f['category']})")
                    st.markdown(f"*{f['message']}*")
                    if f.get("text"):
                        st.caption(f"Context: \"{f['text']}\"")
                    st.markdown("---")
                    
    except Exception as e:
        st.error(f"Failed to process document: {e}")
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        if 'corrected_path' in locals() and os.path.exists(corrected_path):
            os.remove(corrected_path)

else:
    st.info("👈 Please upload a Word file (.docx) on the sidebar to validate it.")
