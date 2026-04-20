import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import fitz  # PyMuPDF
from openai import OpenAI

st.set_page_config(page_title="KDP Ultimate Formatter", layout="wide")

# --- CONFIGURAZIONE AI ---
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception as e:
    st.error("Errore: OPENAI_API_KEY non trovata nei Secrets di Streamlit.")
    st.stop()

# --- FUNZIONI DI SUPPORTO ---

def extract_text_for_ai(file, extension):
    """Estrae il testo per l'analisi AI (fino a 10.000 caratteri)"""
    text = ""
    if extension == "docx":
        doc = Document(file)
        # Prende i paragrafi più significativi
        text = "\n".join([p.text for p in doc.paragraphs if len(p.text) > 20][:100])
    elif extension == "pdf":
        with fitz.open(stream=file.read(), filetype="pdf") as doc:
            for page in doc:
                text += page.get_text()
                if page.number > 20: break
    return text

def advanced_kdp_processing(file):
    """Sitema layout, pulisce Markdown e rimuove spazi vuoti tra capitoli"""
    doc = Document(file)
    
    # 1. Setup Pagina 6x9
    for section in doc.sections:
        section.page_width = Inches(6)
        section.page_height = Inches(9)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.5)

    # 2. Pulizia e Formattazione
    # Usiamo una lista per identificare i paragrafi da rimuovere (quelli vuoti di troppo)
    paragraphs_to_remove = []
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        
        # RIMOZIONE SPAZI VUOTI ECCESSIVI
        # Se il paragrafo è vuoto e quello precedente era vuoto, lo segniamo per la rimozione
        if i > 0 and not text and not doc.paragraphs[i-1].text.strip():
            paragraphs_to_remove.append(para)
            continue

        # Pulizia Markdown e simboli
        clean_text = text.replace("**", "").replace("##", "").replace("###", "").replace("#", "")
        para.text = " ".join(clean_text.split())

        # Gestione Capitoli
        if "CAPITOLO" in para.text.upper() or (len(para.text) < 50 and para.text.isupper()):
            para.style = 'Heading 1'
            para.paragraph_format.page_break_before = True # Forza inizio pagina
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before = Pt(0) # Rimuove spazi vuoti sopra il titolo
            para.paragraph_format.space_after = Pt(24)
        else:
            # Testo Corpo
            if para.text.strip():
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                para.paragraph_format.first_line_indent = Inches(0.2)
                para.paragraph_format.line_spacing = 1.15
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(6)

    # Nota: Rimuovere fisicamente i paragrafi in python-docx è complesso, 
    # quindi settiamo a zero l'altezza dei paragrafi vuoti inutili
    for p in paragraphs_to_remove:
        p.text = ""
        p.paragraph_format.line_spacing = Pt(1)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- INTERFACCIA ---
st.title("🚀 KDP All-in-One: Analisi Profonda & Formattazione")
st.markdown("Carica Word o PDF per ricevere metadati dettagliati e un file Word perfettamente pulito.")

uploaded_file = st.file_uploader("Carica Manoscritto", type=["docx", "pdf"])

if uploaded_file:
    file_ext = uploaded_file.name.split(".")[-1].lower()
    
    col1, col2 = st.columns(2)

    with col1:
        st.header("🔍 Analisi Editoriale AI")
        if st.button("Genera Metadati Dettagliati"):
            with st.spinner("Analisi approfondita del contenuto..."):
                extracted_text = extract_text_for_ai(uploaded_file, file_ext)
                uploaded_file.seek(0)

                prompt = f"""
                Sei un consulente marketing per autori Amazon KDP di alto livello.
                Analizza attentamente questo contenuto:
                ---
                {extracted_text[:9000]}
                ---
                Genera in lingua ITALIANA e in FORMATO TESTO SEMPLICE (ASSOLUTAMENTE NO HTML):

                1. DESCRIZIONE DETTAGLIATA (Minimo 300 parole): 
                   - Un gancio iniziale (hook) potente.
                   - Una spiegazione approfondita di cosa imparerà il lettore o della trama.
                   - Una sezione 'Perché leggere questo libro' con punti elenco (usa il trattino -).
                   - Una call to action finale.

                2. PAROLE CHIAVE SEO (10 frasi): 
                   - Usa parole chiave a 'coda lunga' (es. "come smettere di fumare senza ingrassare" invece di solo "fumo").

                3. PUBBLICO TARGET: Definisci esattamente chi è il lettore ideale.

                4. SUGGERIMENTO PREZZO: Indica un range di prezzo basato sul valore percepito.
                """
                
                resp = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "system", "content": "Sei un esperto di self-publishing professionale."},
                              {"role": "user", "content": prompt}]
                )
                st.subheader("Testi per Amazon (Copia-Incolla)")
                st.text_area("Risultato:", value=resp.choices[0].message.content, height=500)

    with col2:
        st.header("🛠️ Formattazione & Pulizia")
        if file_ext == "docx":
            if st.button("Sistema Layout e Spazi"):
                with st.spinner("Eliminazione spazi vuoti e correzione margini..."):
                    processed_doc = advanced_kdp_processing(uploaded_file)
                    st.success("✅ Documento ottimizzato!")
                    st.download_button(
                        label="⬇️ Scarica Manoscritto KDP",
                        data=processed_doc,
                        file_name=f"KDP_READY_{uploaded_file.name}"
                    )
        else:
            st.info("I file PDF possono essere analizzati dall'AI (Colonna sinistra), ma per la formattazione dei capitoli e degli spazi carica la versione .docx.")
