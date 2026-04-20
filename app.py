import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import fitz
from openai import OpenAI

st.set_page_config(page_title="KDP Impeccable Formatter", layout="wide")

# --- CONFIGURAZIONE AI ---
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception as e:
    st.error("Configura OPENAI_API_KEY nei Secrets di Streamlit.")
    st.stop()

# --- LOGICA DI PULIZIA PROFONDA ---

def delete_paragraph(paragraph):
    """Rimuove fisicamente un paragrafo dalla struttura XML"""
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def impeccable_format(file):
    doc = Document(file)
    
    # 1. Setup Pagina 6x9 (Standard KDP)
    for section in doc.sections:
        section.page_width = Inches(6)
        section.page_height = Inches(9)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)  # Margine interno per rilegatura
        section.right_margin = Inches(0.5) # Margine esterno

    # 2. Rimozione Paragrafi Vuoti e Pulizia Contenuto
    # Iteriamo al contrario per non sballare gli indici durante la rimozione
    for p in list(doc.paragraphs):
        text = p.text.strip()
        
        # Elimina paragrafi completamente vuoti (spazi bianchi tra capitoli)
        if not text:
            delete_paragraph(p)
            continue

        # Pulizia residui Markdown
        clean_text = text.replace("**", "").replace("##", "").replace("###", "").replace("#", "")
        p.text = " ".join(clean_text.split())

        # Gestione Capitoli (Heading 1)
        if "CAPITOLO" in p.text.upper() or (len(p.text) < 40 and p.text.isupper()):
            p.style = 'Heading 1'
            p.paragraph_format.page_break_before = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(30)
        
        # Gestione Indice (ToC Styles)
        elif "INDICE" in p.text.upper() or "SOMMARIO" in p.text.upper():
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(20)
        
        else:
            # Testo Standard
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Inches(0.25)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)

    # 3. Forzatura Font Globale (per evitare indici giganti)
    for style in doc.styles:
        if 'TOC' in style.name: # Stili dell'Indice
            style.font.size = Pt(10)
            style.font.name = 'Georgia'
        if style.name == 'Normal':
            style.font.size = Pt(11)
            style.font.name = 'Georgia'

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- INTERFACCIA ---
st.title("🛡️ KDP Impeccable Formatter")
st.write("Sistemazione chirurgica di spazi, capitoli e indici per Amazon KDP.")

uploaded_file = st.file_uploader("Carica Word o PDF", type=["docx", "pdf"])

if uploaded_file:
    file_ext = uploaded_file.name.split(".")[-1].lower()
    col1, col2 = st.columns(2)

    with col1:
        st.header("✍️ Descrizione & Keyword")
        if st.button("Genera Metadati Professionali"):
            with st.spinner("Analisi testo..."):
                # Estrazione testo
                raw_text = ""
                if file_ext == "docx":
                    d = Document(uploaded_file)
                    raw_text = "\n".join([p.text for p in d.paragraphs[:100]])
                else:
                    with fitz.open(stream=uploaded_file.read(), filetype="pdf") as pdf:
                        for page in pdf[:15]: raw_text += page.get_text()
                uploaded_file.seek(0)

                prompt = f"""
                Analizza questo libro e genera esclusivamente in TESTO SEMPLICE (NO HTML):
                1. DESCRIZIONE DETTAGLIATA (400+ parole): Inizia con un gancio emotivo, spiega i capitoli principali e chiudi con una call to action.
                2. 10 KEYWORD A CODA LUNGA: Frasi specifiche che i lettori cercano su Amazon.
                3. CATEGORIE KDP: Suggerisci le 3 migliori.
                
                NON includere il pubblico target. Rispondi in Italiano.
                Testo: {raw_text[:8000]}
                """
                
                resp = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": prompt}]
                )
                st.text_area("Copia da qui:", value=resp.choices[0].message.content, height=500)

    with col2:
        st.header("⚙️ Formattazione Word")
        if file_ext == "docx":
            if st.button("Esegui Sistemazione Impeccabile"):
                with st.spinner("Rimozione spazi e calibrazione indice..."):
                    clean_file = impeccable_format(uploaded_file)
                    st.success("✅ Documento ripulito e calibrato!")
                    st.download_button("⬇️ Scarica File 6x9", clean_file, f"KDP_PERFECT_{uploaded_file.name}")
        else:
            st.warning("La formattazione chirurgica richiede un file .docx.")
