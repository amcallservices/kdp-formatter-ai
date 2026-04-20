import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, ns
import io
import fitz
from openai import OpenAI

st.set_page_config(page_title="KDP Professional Suite", layout="wide")

try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    st.error("Configura OPENAI_API_KEY nei Secrets di Streamlit.")
    st.stop()

# --- FUNZIONI DI SUPPORTO TECNICO ---

def delete_paragraph(paragraph):
    """Rimuove fisicamente il paragrafo dalla struttura XML del documento."""
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def add_page_numbers(doc):
    """Aggiunge la numerazione automatica centrata nel piè di pagina."""
    for section in doc.sections:
        footer = section.footer
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Creazione del campo dinamico 'PAGE'
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(ns.qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.text = "PAGE"
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(ns.qn('w:fldCharType'), 'end')
        
        run = p.add_run()
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)

def impeccable_format(file):
    doc = Document(file)
    
    # 1. Configurazione Pagina 6x9
    for section in doc.sections:
        section.page_width = Inches(6)
        section.page_height = Inches(9)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8) # Gutter (Margine rilegatura)
        section.right_margin = Inches(0.5)

    # 2. Pulizia Totale e Formattazione
    for p in list(doc.paragraphs):
        text = p.text.strip()
        
        # Eliminazione chirurgica dei paragrafi vuoti
        if not text:
            delete_paragraph(p)
            continue

        # Pulizia residui Markdown e correzione spazi
        clean_text = text.replace("**", "").replace("##", "").replace("#", "")
        p.text = " ".join(clean_text.split())

        # Riconoscimento Titoli (Heading 1)
        if len(p.text) < 60 and ("CAPITOLO" in p.text.upper() or p.text.isupper()):
            p.style = 'Heading 1'
            p.paragraph_format.page_break_before = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(30)
        else:
            # Testo Corpo
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Inches(0.25)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(6)

    # 3. Scala Indice e Numerazione
    for style in doc.styles:
        if 'TOC' in style.name:
            style.font.size = Pt(10) # Indice compatto
        if style.name == 'Normal':
            style.font.name = 'Georgia'
            style.font.size = Pt(11)

    add_page_numbers(doc)
    
    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- INTERFACCIA ---
st.title("🛡️ KDP Professional Suite")

uploaded_file = st.file_uploader("Carica Manoscritto", type=["docx", "pdf"])

if uploaded_file:
    file_ext = uploaded_file.name.split(".")[-1].lower()
    col1, col2 = st.columns(2)

    with col1:
        st.header("✍️ AI Marketing Copy")
        if st.button("Genera Metadati Dettagliati"):
            with st.spinner("Analisi approfondita del contenuto..."):
                # Estrazione testo per contesto AI
                context = ""
                if file_ext == "docx":
                    d = Document(uploaded_file)
                    context = "\n".join([p.text for p in d.paragraphs[:100]])
                else:
                    with fitz.open(stream=uploaded_file.read(), filetype="pdf") as pdf:
                        for page in pdf[:15]: context += page.get_text()
                uploaded_file.seek(0)

                prompt = f"""
                Analizza questo libro: {context[:8000]}
                
                Genera in lingua ITALIANA e in TESTO SEMPLICE (ASSOLUTAMENTE NO HTML):
                
                1. DESCRIZIONE MARKETING (400-500 parole):
                   - HOOK: Una frase d'apertura scioccante o emozionante.
                   - IL PROBLEMA: Descrivi la frustrazione del lettore.
                   - LA SOLUZIONE: Presenta il libro come la guida definitiva.
                   - COSA IMPARERAI: Lista dettagliata di benefici (usa '-').
                   - CALL TO ACTION: Invito all'acquisto potente.
                   (NON aggiungere target audience).

                2. 10 KEYWORD A CODA LUNGA: 
                   Fornisci 10 frasi specifiche (3-5 parole l'una) che un utente cercherebbe su Amazon per trovare questo libro.
                """
                
                resp = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": prompt}]
                )
                st.text_area("Copia per KDP:", value=resp.choices[0].message.content, height=500)

    with col2:
        st.header("⚙️ Formattazione Impeccabile")
        if file_ext == "docx":
            if st.button("Esegui Sistemazione Finale"):
                with st.spinner("Pulizia XML e numerazione pagine..."):
                    clean_file = impeccable_format(uploaded_file)
                    st.success("✅ Layout sistemato, spazi eliminati e pagine numerate!")
                    st.download_button("⬇️ Scarica File 6x9", clean_file, f"KDP_FINAL_{uploaded_file.name}")
