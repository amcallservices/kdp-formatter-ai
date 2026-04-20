import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import fitz  # PyMuPDF
from openai import OpenAI

st.set_page_config(page_title="KDP Pro Formatter", page_icon="📚", layout="wide")

# --- CONFIGURAZIONE API ---
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception as e:
    st.error("Chiave API non configurata nei Secrets di Streamlit.")
    st.stop()

# --- FUNZIONI DI ESTREZIONE E FORMATTAZIONE ---

def extract_text(file, extension):
    text = ""
    if extension == "docx":
        doc = Document(file)
        text = "\n".join([p.text for p in doc.paragraphs])
    elif extension == "pdf":
        with fitz.open(stream=file.read(), filetype="pdf") as doc:
            for page in doc:
                text += page.get_text()
                if page.number > 15: break # Limite per contesto AI
    return text

def format_entire_document(file):
    doc = Document(file)
    
    # 1. Impostazione globale Formato 6x9 e Margini per tutte le sezioni
    for section in doc.sections:
        section.page_width = Inches(6)
        section.page_height = Inches(9)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)   # Margine interno (Gutter)
        section.right_margin = Inches(0.5)  # Margine esterno
        
    # 2. Ottimizzazione stili testo per l'intero contenuto
    for paragraph in doc.paragraphs:
        # Giustifica il testo (standard per i libri)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        # Imposta interlinea singola o 1.15 per leggibilità
        paragraph.paragraph_format.line_spacing = 1.15
        
    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- INTERFACCIA UTENTE ---

st.title("📚 KDP Formatter & AI Metadata Expert")
st.info("Carica il tuo manoscritto. L'AI genererà i testi per Amazon e il sistema formatterà l'intero documento (inclusi margini e layout dell'indice).")

uploaded_file = st.file_uploader("Trascina qui il tuo file (DOCX o PDF)", type=["docx", "pdf"])

if uploaded_file:
    file_ext = uploaded_file.name.split(".")[-1].lower()
    
    col_ai, col_format = st.columns(2)

    with col_ai:
        st.header("✍️ Descrizione e Parole Chiave")
        if st.button("Analizza e Genera Testi"):
            with st.spinner("L'AI sta analizzando il contenuto..."):
                full_text = extract_text(uploaded_file, file_ext)
                uploaded_file.seek(0)

                prompt = f"""
                Sei un esperto SEO di Amazon KDP. Analizza questo testo:
                {full_text[:7000]}
                
                Genera:
                1. DESCRIZIONE: Un testo persuasivo per la vendita con tag HTML (<b>, <i>, <ul>).
                2. PAROLE CHIAVE: 7 frasi chiave SEO separate da virgola.
                3. CATEGORIE: 2 categorie suggerite per KDP.
                Rispondi in Italiano.
                """

                response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": prompt}]
                )
                
                result = response.choices[0].message.content
                st.success("Testi generati con successo!")
                
                # Visualizzazione testuale pronta per il copia-incolla
                st.subheader("Copia da qui per Amazon KDP:")
                st.code(result, language="html") 

    with col_format:
        st.header("📄 Formattazione Integrale")
        if file_ext == "docx":
            st.write("Configurazione: **6x9 pollici**, margini ottimizzati per la rilegatura.")
            if st.button("Formatta Intero Libro"):
                with st.spinner("Rielaborazione layout in corso..."):
                    formatted_docx = format_entire_document(uploaded_file)
                    uploaded_file.seek(0)
                    
                    st.success("✅ Formattazione completata!")
                    st.warning("⚠️ Nota: Quando apri il file in Word, clicca col tasto destro sull'Indice e seleziona 'Aggiorna campo -> Aggiorna intero sommario' per allineare i numeri di pagina.")
                    
                    st.download_button(
                        label="⬇️ Scarica File Pronto per KDP",
                        data=formatted_docx,
                        file_name=f"KDP_READY_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
        else:
            st.error("La formattazione automatica richiede un file .docx. Per i PDF posso generare solo i metadati AI.")

st.sidebar.markdown("""
### Istruzioni rapide:
1. **Carica** il tuo file.
2. **Genera i metadati**: Copia il codice HTML risultante nella sezione 'Descrizione' di Amazon KDP.
3. **Formatta**: Scarica il file Word e caricalo nella sezione 'Contenuto Manoscritto' di KDP.
""")
