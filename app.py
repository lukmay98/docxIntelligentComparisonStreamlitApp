import streamlit as st
import docx
import difflib

# Seite konfigurieren
st.set_page_config(page_title="Docx Vergleicher", layout="wide")
st.title("📄 Intelligenter Word-Dokumentenvergleich")
st.write("Vergleiche zwei Versionen einer .docx-Datei. Der Algorithmus erkennt Einfügungen und Löschungen, ohne den Rhythmus des restlichen Dokuments zu verlieren.")

# Funktion zum Extrahieren von Text aus Word-Dateien
def read_docx(file):
    doc = docx.Document(file)
    # Extrahiere jeden Absatz und behalte die Zeilenstruktur bei
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return full_text

# Upload-Bereich für zwei (oder mehr) Dateien
col1, col2 = st.columns(2)
with col1:
    file1 = st.file_uploader("Original-Dokument (Version 1)", type=["docx"])
with col2:
    file2 = st.file_uploader("Geändertes Dokument (Version 2)", type=["docx"])

if file1 and file2:
    with st.spinner("Lese Dokumente und vergleiche intelligent..."):
        # Text extrahieren
        text1 = read_docx(file1)
        text2 = read_docx(file2)
        
        # HtmlDiff generiert eine intelligente Side-by-Side Vergleichstabelle
        # Es vergleicht Zeilen UND erkennt Unterschiede innerhalb der Zeile!
        differ = difflib.HtmlDiff()
        
        # Erstelle die HTML-Tabelle
        html_diff = differ.make_table(
            text1, 
            text2, 
            fromdesc=file1.name, 
            todesc=file2.name,
            context=True,      # Zeigt nur geänderte Absätze mit etwas Kontext darum
            numlines=2         # Anzahl der Kontextzeilen
        )
        
        st.success("Vergleich erfolgreich abgeschlossen!")
        
        # Ein wenig CSS, um die difflib-Tabelle an Streamlit anzupassen
        custom_css = """
        <style>
            table.diff {font-family: Arial, sans-serif; width: 100%; border-collapse: collapse;}
            table.diff th {background-color: #f0f2f6; padding: 8px;}
            table.diff td {padding: 6px; border: 1px solid #ddd; vertical-align: top;}
            .diff_add {background-color: #d4edda; color: #155724;} /* Grün für Neues */
            .diff_sub {background-color: #f8d7da; color: #721c24;} /* Rot für Gelöschtes */
            .diff_chg {background-color: #fff3cd; color: #856404;} /* Gelb für Änderungen */
        </style>
        """
        
        # HTML-Ausgabe direkt in der Streamlit-App rendern
        st.components.v1.html(custom_css + html_diff, height=700, scrolling=True)
else:
    st.info("Bitte lade auf beiden Seiten eine .docx-Datei hoch, um den Vergleich zu starten.")