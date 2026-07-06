import streamlit as st
import docx
from docx.enum.text import WD_COLOR_INDEX
from difflib import SequenceMatcher
import io

# -- UI KONFIGURATION --
st.set_page_config(page_title="DocFusion Pro", layout="wide")

st.markdown("""
    <style>
    .diff-box { padding: 10px; border-radius: 5px; border: 1px solid #ddd; height: 400px; overflow-y: auto; font-family: monospace; white-space: pre-wrap; }
    .add { background-color: #d4edda; text-decoration: underline; }
    .rem { background-color: #f8d7da; text-decoration: line-through; }
    </style>
""", unsafe_allow_html=True)

# -- FUNKTIONEN --
def reset_app():
    for key in st.session_state.keys():
        del st.session_state[key]

def get_text(file):
    doc = docx.Document(file)
    return [p.text for p in doc.paragraphs if p.text.strip() != ""]

def create_fused_docx(base_text, new_text):
    matcher = SequenceMatcher(None, base_text, new_text)
    doc = docx.Document()
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            for line in base_text[i1:i2]:
                doc.add_paragraph(line)
        elif tag == 'insert':
            for line in new_text[j1:j2]:
                p = doc.add_paragraph()
                run = p.add_run(line)
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        elif tag == 'replace':
            # Altes löschen (optional), Neues einfügen mit Markierung
            for line in new_text[j1:j2]:
                p = doc.add_paragraph()
                run = p.add_run(f"[EINGEFÜGT]: {line}")
                run.font.highlight_color = WD_COLOR_INDEX.TURQUOISE
                
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

# -- APP LAYOUT --
st.title("📑 Intelligent Document Fusion & Comparison")

if 'started' not in st.session_state:
    col1, col2 = st.columns(2)
    with col1:
        f1 = st.file_uploader("Original (V1)", type="docx", key="u1")
    with col2:
        f2 = st.file_uploader("Überarbeitung (V2)", type="docx", key="u2")
    
    if f1 and f2:
        t1 = get_text(f1)
        t2 = get_text(f2)
        
        # Side-by-Side Ansicht mit Wrapping
        st.subheader("Detail-Analyse (Intelligente Synchronisation)")
        c1, c2 = st.columns(2)
        
        matcher = SequenceMatcher(None, t1, t2)
        diff_html_1 = ""
        diff_html_2 = ""
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                chunk = "\n".join(t1[i1:i2]) + "\n"
                diff_html_1 += chunk
                diff_html_2 += chunk
            elif tag == 'insert':
                diff_html_2 += f'<span class="add">{" ".join(t2[j1:j2])}</span>\n'
            elif tag == 'delete':
                diff_html_1 += f'<span class="rem">{" ".join(t1[i1:i2])}</span>\n'
            elif tag == 'replace':
                diff_html_1 += f'<span class="rem">{" ".join(t1[i1:i2])}</span>\n'
                diff_html_2 += f'<span class="add">{" ".join(t2[j1:j2])}</span>\n'

        c1.markdown(f'<div class="diff-box">{diff_html_1}</div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="diff-box">{diff_html_2}</div>', unsafe_allow_html=True)
        
        # Fusion & Download
        st.divider()
        fused_file = create_fused_docx(t1, t2)
        st.download_button("Fusioniertes Dokument herunterladen (.docx)", 
                           data=fused_file, 
                           file_name="fusion_result.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

st.button("Neuen Vergleich starten", on_click=reset_app, type="primary")
