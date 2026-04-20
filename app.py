import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
from openai import OpenAI

# Configurazione API
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    st.error("API Key mancante nei Secrets.")
    st.stop()

def impeccabile_format(file):
    doc = Document(file)
    
    # 1. Setup Pagina 6x9 (Standard KDP)
    for section in doc.sections:
        section.page_width = Inches(6)
        section.page_height = Inches(9)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.5)

    # 2. ELIMINAZIONE REALE PARAGRAFI VUOTI
    # Creiamo una lista dei paragrafi da eliminare per non alterare l'indice durante il ciclo
    paragraphs = list(doc.paragraphs)
    for p in paragraphs:
        # Se il paragrafo è vuoto o contiene solo spazi/tab, lo rimuoviamo dall'XML
        if not p.text.strip():
            p_element = p._element
            p_element.getparent().remove(p_element)
            continue

        # 3. RILEVAMENTO TITOLI MIGLIORATO
        text = p.text.strip()
        # Pulizia residui Markdown
        text = text.replace("**", "").replace("##", "").replace("#", "")
        p.text = " ".join(text.split()) # Rimuove doppi spazi interni

        # Un titolo vero è solitamente corto (<60 caratteri)
        is_short = len(p.text) < 60
        is_chapter = "CAPITOLO" in p.text.upper() or "PARTE" in p.text.upper()
        is_intro = p.text.upper() in ["PREFAZIONE", "INTRODUZIONE", "RINGRAZIAMENTI", "INDICE", "SOMMARIO"]

        if is_short and (is_chapter or is_intro or p.text.isupper()):
            # --- FORMATTAZIONE TITOLO ---
            p.style = 'Heading 1'
            p.paragraph_format.page_break_before = True # Forza nuova pagina
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(30)
        
        elif is_short and re.match(r'^\d+(\.\d+)*\s', p.text):
            # Sottotitoli tipo 1.1, 1.2
            p.style = 'Heading 2'
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.page_break_before = False
        
        else:
            # --- FORMATTAZIONE CORPO TESTO (Standard) ---
            p.style = 'Normal'
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Inches(0.25)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)

    # 4. CALIBRAZIONE FONT
    for style in doc.styles:
        if 'TOC' in style.name: # Indice
            style.font.size = Pt(10)
        if style.name == 'Normal':
            style.font.size = Pt(11)
            style.font.name = 'Georgia'

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- UI STREAMLIT ---
st.title("📚 KDP Formatter PRO (Fix Edizione)")

uploaded_file = st.file_uploader("Carica il file Word", type=["docx"])

if uploaded_file:
    if st.button("Esegui Pulizia Totale"):
        with st.spinner("Riparazione struttura in corso..."):
            file_pronto = impeccabile_format(uploaded_file)
            st.success("Sistemazione completata: Vuoti eliminati e titoli calibrati.")
            st.download_button("Scarica File 6x9", file_pronto, f"KDP_FIXED_{uploaded_file.name}")
