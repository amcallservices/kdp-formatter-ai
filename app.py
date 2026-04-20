import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import fitz
from openai import OpenAI

st.set_page_config(page_title="KDP Impeccable Formatter", layout="wide")

try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception as e:
    st.error("Configura OPENAI_API_KEY nei Secrets di Streamlit.")
    st.stop()

# --- FUNZIONE NUCLEARE PER ELIMINARE GLI SPAZI ---
def delete_paragraph(paragraph):
    """Questa funzione non 'nasconde' il paragrafo, ma lo elimina fisicamente dal file."""
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def impeccable_format(file):
    doc = Document(file)
    
    # 1. IMPOSTAZIONE PAGINA 6x9 (Dimensioni KDP)
    for section in doc.sections:
        section.page_width = Inches(6)
        section.page_height = Inches(9)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)  # Margine interno per la colla
        section.right_margin = Inches(0.5) 

    # 2. PULIZIA CHIRURGICA E STILI
    # Usiamo una lista fissa di paragrafi per poterli eliminare in sicurezza durante il ciclo
    for p in list(doc.paragraphs):
        text = p.text.strip()
        
        # Elimina immediatamente le righe vuote e i "buchi" tra i capitoli
        if not text:
            delete_paragraph(p)
            continue

        # Distrugge i residui del formato Markdown (##, **)
        clean_text = text.replace("**", "").replace("##", "").replace("###", "").replace("#", "")
        p.text = " ".join(clean_text.split())

        testo_maiuscolo = p.text.upper()

        # Logica per identificare i CAPITOLI e le PARTI
        if "CAPITOLO" in testo_maiuscolo or "PARTE" in testo_maiuscolo or "PREFAZIONE" in testo_maiuscolo or "RINGRAZIAMENTI" in testo_maiuscolo:
            p.style = 'Heading 1'
            p.paragraph_format.page_break_before = True # Forza su nuova pagina
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)   # Zero spazio sopra
            p.paragraph_format.space_after = Pt(24)   # Spazio solo sotto
        
        # Sottotitoli (es: 1.1 La prigione...)
        elif len(p.text) < 50 and p.text[0].isdigit():
            p.style = 'Heading 2'
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)

        # Gestione voci dell'INDICE (se presenti come testo semplice)
        elif "INDICE" in testo_maiuscolo or "SOMMARIO" in testo_maiuscolo:
            p.style = 'Heading 1'
            p.paragraph_format.page_break_before = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(20)
        
        # Testo del Libro
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Inches(0.25)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)

    # 3. RIDIMENSIONAMENTO AUTOMATICO INDICE (TOC) E FONT GLOBALE
    # Forza un font pulito e riduce la dimensione dell'indice per non sbordare nel 6x9
    for style in doc.styles:
        if 'TOC' in style.name: 
            style.font.size = Pt(9)
        if style.name == 'Normal':
            style.font.size = Pt(11)

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- INTERFACCIA ---
st.title("🛡️ KDP Impeccable Formatter")
st.write("Sistemazione chirurgica di spazi, layout e metadati. Nessun artefatto ammesso.")

uploaded_file = st.file_uploader("Carica Manoscritto (.docx o .pdf)", type=["docx", "pdf"])

if uploaded_file:
    file_ext = uploaded_file.name.split(".")[-1].lower()
    col1, col2 = st.columns(2)

    with col1:
        st.header("✍️ Copywriting KDP")
        if st.button("Genera Solo Descrizione e Keyword"):
            with st.spinner("Creazione testi persuasivi..."):
                testo_grezzo = ""
                if file_ext == "docx":
                    d = Document(uploaded_file)
                    testo_grezzo = "\n".join([p.text for p in d.paragraphs[:100]])
                else:
                    with fitz.open(stream=uploaded_file.read(), filetype="pdf") as pdf:
                        for page in pdf[:15]: testo_grezzo += page.get_text()
                uploaded_file.seek(0)

                prompt = f"""
                Analizza questo libro: {testo_grezzo[:8000]}
                
                Genera in lingua ITALIANA e in TESTO SEMPLICE (ASSOLUTAMENTE NO HTML, NO BOLD):
                
                1. DESCRIZIONE DEL LIBRO (300+ parole): Parti subito forte con un'introduzione che colpisce. Spiega esattamente cosa scoprirà il lettore. Inserisci un elenco puntato (usa il simbolo '-') sui benefici principali. Termina con un invito all'acquisto (es. 'Scorri verso l'alto e prendi la tua copia'). Non scrivere chi è il target.
                
                2. 10 KEYWORD A CODA LUNGA: Fornisci solo le 10 frasi separate da virgola (es. come smettere di fumare facilmente, metodo per smettere di fumare, ecc.).
                """
                
                resp = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": prompt}]
                )
                st.text_area("Copia da qui e incolla su Amazon:", value=resp.choices[0].message.content, height=500)

    with col2:
        st.header("⚙️ Layout Impeccabile")
        if file_ext == "docx":
            if st.button("Distruggi Spazi Vuoti e Formatta"):
                with st.spinner("Eliminazione XML degli spazi e calibrazione margini..."):
                    clean_file = impeccable_format(uploaded_file)
                    st.success("✅ Layout chirurgicamente pulito!")
                    st.download_button("⬇️ Scarica File 6x9 Pronto", clean_file, f"KDP_PERFETTO_{uploaded_file.name}")
        else:
            st.warning("Per applicare il formato impeccabile serve la versione in Word (.docx).")
