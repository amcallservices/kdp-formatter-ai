import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import fitz # PyMuPDF per i PDF
from openai import OpenAI

st.set_page_config(page_title="KDP Content & Layout Master", layout="wide")

# --- CONFIGURAZIONE AI ---
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception as e:
    st.error("Configura OPENAI_API_KEY nei Secrets di Streamlit.")
    st.stop()

# --- FUNZIONE DI SISTEMAZIONE CONTENUTO E LAYOUT ---
def process_kdp_book(file):
    doc = Document(file)
    
    # 1. SETTING PAGINA (6x9 pollici - Standard KDP)
    for section in doc.sections:
        section.page_width = Inches(6)
        section.page_height = Inches(9)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8) # Margine interno (Gutter)
        section.right_margin = Inches(0.5)

    # 2. SISTEMAZIONE CONTENUTO (LOGICA EDITORIALE)
    for para in doc.paragraphs:
        # Pulizia spazi bianchi extra
        para.text = " ".join(para.text.split())
        
        # Riconoscimento Capitoli e Formattazione
        # Se la riga è "Capitolo X" o è molto corta e in maiuscolo, la trattiamo come titolo
        if "CAPITOLO" in para.text.upper() or (para.text.isupper() and 2 < len(para.text) < 50):
            para.style = 'Heading 1'
            para.paragraph_format.page_break_before = True # Inizia su nuova pagina
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            # Testo standard: Giustificato con rientro
            if para.text.strip():
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                para.paragraph_format.first_line_indent = Inches(0.2)
                para.paragraph_format.line_spacing = 1.15

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- INTERFACCIA ---
st.title("📖 KDP Master: Formattazione & Metadati Testuali")
st.write("Carica il tuo file per sistemare il layout del libro e generare i testi (senza HTML) per Amazon.")

uploaded_file = st.file_uploader("Carica Manoscritto (Word o PDF)", type=["docx", "pdf"])

if uploaded_file:
    file_ext = uploaded_file.name.split(".")[-1].lower()
    
    col1, col2 = st.columns(2)

    # --- COLONNA 1: GENERAZIONE AI (TESTO SEMPLICE) ---
    with col1:
        st.header("🤖 Descrizione e Keyword (Testo)")
        if st.button("Genera Testi per Pubblicazione"):
            with st.spinner("Analisi in corso..."):
                # Estrazione testo
                text_content = ""
                if file_ext == "docx":
                    temp_doc = Document(uploaded_file)
                    text_content = "\n".join([p.text for p in temp_doc.paragraphs[:60]])
                else:
                    with fitz.open(stream=uploaded_file.read(), filetype="pdf") as d:
                        for p in d[:10]: text_content += p.get_text()
                uploaded_file.seek(0)

                # PROMPT SENZA HTML
                prompt = f"""
                Sei un esperto di marketing Amazon KDP. Analizza questo estratto:
                {text_content[:6000]}
                
                Genera i seguenti contenuti in FORMATO TESTO SEMPLICE (NON USARE TAG HTML come <b>, <i> o <ul>):
                1) DESCRIZIONE: Scrivi una descrizione persuasiva e professionale. Usa solo paragrafi e, se necessario, elenchi puntati usando il trattino (-).
                2) PAROLE CHIAVE: Una lista di 7 frasi chiave SEO separate da virgola.
                3) CATEGORIE: Suggerisci le 2 migliori categorie Amazon.
                
                Rispondi esclusivamente in lingua Italiana.
                """
                
                resp = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": prompt}]
                )
                
                testo_finale = resp.choices[0].message.content
                
                st.success("✅ Metadati Generati!")
                # Visualizzazione pulita senza formato codice o HTML
                st.text_area("Copia la Descrizione e le Keyword da qui:", value=testo_finale, height=400)

    # --- COLONNA 2: SISTEMAZIONE FILE ---
    with col2:
        st.header("🛠️ Sistemazione File Word")
        if file_ext == "docx":
            st.write("Configurazione: **6x9 pollici**, giustificazione testo e gestione capitoli.")
            if st.button("Sitema e Formatta Contenuto"):
                with st.spinner("Sistemazione in corso..."):
                    final_docx = process_kdp_book(uploaded_file)
                    st.success("✅ Documento sistemato!")
                    st.download_button(
                        label="⬇️ Scarica Libro Pronto", 
                        data=final_docx, 
                        file_name=f"KDP_PRONTO_{uploaded_file.name}"
                    )
        else:
            st.warning("⚠️ Per la sistemazione del contenuto è necessario un file Word (.docx).")
