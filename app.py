import streamlit as st
from docx import Document
from docx.shared import Inches
import io
import fitz  # PyMuPDF per leggere i PDF
from openai import OpenAI

# Configurazione della pagina
st.set_page_config(
    page_title="KDP Master Tool", 
    page_icon="📚", 
    layout="wide"
)

# --- CONFIGURAZIONE API OPENAI ---
# Recupera la chiave dai Secrets di Streamlit
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception as e:
    st.error("⚠️ Errore: OpenAI API Key non trovata. Vai nei Settings di Streamlit -> Secrets e aggiungi OPENAI_API_KEY.")
    st.stop()

# --- FUNZIONI TECNICHE ---

def extract_text(file, extension):
    """Estrae il testo dal file per darlo in pasto all'AI"""
    text = ""
    if extension == "docx":
        doc = Document(file)
        text = "\n".join([p.text for p in doc.paragraphs])
    elif extension == "pdf":
        with fitz.open(stream=file.read(), filetype="pdf") as doc:
            # Leggiamo le prime 15 pagine per il contesto
            for page in doc:
                text += page.get_text()
                if doc.page_count > 15 and page.number > 15:
                    break
    return text

def format_docx_kdp(file):
    """Formatta il file Word con le specifiche KDP 6x9 pollici"""
    doc = Document(file)
    for section in doc.sections:
        # Formato 6x9 pollici
        section.page_width = Inches(6)
        section.page_height = Inches(9)
        
        # Margini professionali KDP
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)  # Gutter (margine interno rilegatura)
        section.right_margin = Inches(0.5) # Margine esterno
        
    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- INTERFACCIA UTENTE ---

st.title("📚 KDP All-in-One: Formattazione & AI")
st.markdown("""
Questo strumento analizza il tuo libro (Word o PDF), genera i metadati per Amazon (Descrizione e Parole Chiave) 
e formatta automaticamente il file Word per la stampa in **6x9 pollici**.
""")

st.divider()

# Caricamento file
uploaded_file = st.file_uploader("Carica il tuo manoscritto (DOCX o PDF)", type=["docx", "pdf"])

if uploaded_file:
    file_ext = uploaded_file.name.split(".")[-1].lower()
    
    # Creiamo due colonne per l'interfaccia
    col_ai, col_format = st.columns(2)

    # 1. GENERAZIONE METADATI CON AI
    with col_ai:
        st.header("🤖 Assistente Pubblicazione AI")
        if st.button("Genera Metadati KDP"):
            with st.spinner("L'AI sta leggendo il tuo libro..."):
                try:
                    # Estrazione testo
                    content = extract_text(uploaded_file, file_ext)
                    # Ripristina il file per l'uso successivo
                    uploaded_file.seek(0)
                    
                    # Prompt ottimizzato per KDP
                    prompt = f"""
                    Sei un esperto di marketing Amazon KDP. Analizza il seguente testo estratto dal libro:
                    ---
                    {content[:8000]} 
                    ---
                    Genera:
                    1. Una DESCRIZIONE libro persuasiva in formato HTML (usa <b>, <i>, <ul>, <li>).
                    2. Una lista di 7 PAROLE CHIAVE (keyword phrases) ottimizzate per la SEO di Amazon.
                    3. Suggerisci la CATEGORIA KDP più adatta.
                    Rispondi in lingua ITALIANA.
                    """
                    
                    response = client.chat.completions.create(
                        model="gpt-4o",
                        messages=[{"role": "system", "content": "Sei un esperto di self-publishing."},
                                  {"role": "user", "content": prompt}]
                    )
                    
                    st.success("✅ Analisi Completata!")
                    st.markdown(response.choices[0].message.content)
                    
                except Exception as e:
                    st.error(f"Errore durante l'analisi AI: {e}")

    # 2. FORMATTAZIONE FILE
    with col_format:
        st.header("📄 Formattazione Manoscritto")
        if file_ext == "docx":
            st.info("Formato rilevato: Word. Posso formattare i margini per il cartaceo 6x9.")
            if st.button("Applica Formattazione KDP"):
                with st.spinner("Formattazione in corso..."):
                    # Processo di formattazione
                    formatted_file = format_docx_kdp(uploaded_file)
                    uploaded_file.seek(0)
                    
                    st.balloons()
                    st.download_button(
                        label="⬇️ Scarica Word Formattato (6x9)",
                        data=formatted_file,
                        file_name=f"KDP_6x9_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
        else:
            st.warning("⚠️ Nota: La formattazione automatica dei margini è disponibile solo per i file .docx.")
            st.write("I file PDF sono 'fissi'. Per cambiare i margini di un PDF, devi modificare il file Word originale e ricaricarlo qui.")

st.sidebar.markdown("---")
st.sidebar.info("Sviluppato per autori Amazon KDP")
