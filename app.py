import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import fitz
from openai import OpenAI

st.set_page_config(page_title="KDP Content Editor PRO", layout="wide")

# --- CONFIGURAZIONE AI ---
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception as e:
    st.error("Configura OPENAI_API_KEY nei Secrets di Streamlit.")
    st.stop()

# --- FUNZIONE AVANZATA DI SISTEMAZIONE CONTENUTO ---
def clean_and_format_kdp(file):
    doc = Document(file)
    
    # 1. Impostazione Pagina 6x9
    for section in doc.sections:
        section.page_width = Inches(6)
        section.page_height = Inches(9)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.5)

    paragraphs = doc.paragraphs
    prev_text = ""

    # Usiamo un ciclo per poter rimuovere elementi se necessario
    for i in range(len(paragraphs)):
        p = paragraphs[i]
        original_text = p.text.strip()
        
        # A. RIMOZIONE ARTEFATTI MARKDOWN E PULIZIA SPAZI
        # Rimuove **, ##, ### e spazi doppi
        clean_text = original_text.replace("**", "").replace("##", "").replace("###", "")
        clean_text = " ".join(clean_text.split())
        p.text = clean_text

        # B. ELIMINAZIONE TITOLI DOPPI / RIDONDANTI
        # Se il testo è quasi uguale al precedente (es. "Capitolo 1" e poi "1. Capitolo 1"), svuota il secondo
        if clean_text.upper() == prev_text.upper() and len(clean_text) > 2:
            p.text = ""
            continue
        
        # C. LOGICA DI STILE (CAPITOLI E SOTTOTITOLI)
        if "CAPITOLO" in clean_text.upper() or (original_text.startswith("#") and len(clean_text) < 60):
            p.style = 'Heading 1'
            p.paragraph_format.page_break_before = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(18)
        elif len(clean_text) < 50 and (original_text.startswith("##") or clean_text[0:3].replace(".","").isdigit()):
            # Riconosce sottotitoli come 1.1, 1.2 o righe brevi
            p.style = 'Heading 2'
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(12)
        else:
            # D. TESTO CORPO (GIUSTIFICATO)
            if len(clean_text) > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Inches(0.2)
                p.paragraph_format.line_spacing = 1.15
        
        if clean_text:
            prev_text = clean_text

    # Rimuove i paragrafi rimasti vuoti dopo la pulizia
    for p in doc.paragraphs:
        if not p.text.strip() and p.style.name != 'Heading 1':
            # Nota: python-docx non elimina facilmente, ma possiamo ridurne l'ingombro
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- INTERFACCIA STREAMLIT ---
st.title("📖 Editor Editoriale KDP")
st.markdown("Questo strumento pulisce il codice Markdown, elimina i titoli doppi e formatta il libro per la stampa professionale.")

uploaded_file = st.file_uploader("Carica il file .docx", type=["docx"])

if uploaded_file:
    col1, col2 = st.columns(2)

    with col1:
        st.header("🤖 Metadati (Solo Testo)")
        if st.button("Genera Descrizione e Keyword"):
            with st.spinner("L'AI sta analizzando il manoscritto..."):
                doc_temp = Document(uploaded_file)
                text_sample = "\n".join([p.text for p in doc_temp.paragraphs[:40]])
                uploaded_file.seek(0)

                prompt = f"""
                Analizza questo libro: {text_sample[:5000]}
                Genera in lingua ITALIANA e in FORMATO TESTO SEMPLICE (ASSOLUTAMENTE NO HTML, NO TAG):
                1) DESCRIZIONE: Un testo di vendita avvincente.
                2) KEYWORDS: 7 parole chiave separate da virgola.
                3) CATEGORIE: 2 categorie KDP.
                """
                
                resp = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": prompt}]
                )
                st.text_area("Copia per KDP:", value=resp.choices[0].message.content, height=350)

    with col2:
        st.header("🛠️ Pulizia Contenuto")
        if st.button("Sitema Contenuto e Layout"):
            with st.spinner("Riparazione testo in corso..."):
                cleaned_docx = clean_and_format_kdp(uploaded_file)
                st.success("✅ Pulizia completata! Simboli Markdown rimossi e titoli sistemati.")
                st.download_button(
                    label="⬇️ Scarica Manoscritto Pulito",
                    data=cleaned_docx,
                    file_name=f"KDP_CLEAN_{uploaded_file.name}"
                )
