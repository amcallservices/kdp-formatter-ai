import streamlit as st
from docx import Document
from docx.shared import Inches
import io
from openai import OpenAI

st.set_page_config(page_title="KDP Formatter & AI Assistant", page_icon="📖")

st.title("📖 KDP Formatter & OpenAI Assistant")

# --- BARRA LATERALE: CONFIGURAZIONE ---
st.sidebar.header("Impostazioni")
api_key = st.sidebar.text_input("Inserisci OpenAI API Key:", type="password")

# --- SEZIONE 1: AI GENERATOR (OPENAI) ---
st.header("1. Generatore Metadati AI")
st.write("Inserisci i dettagli del tuo libro per generare descrizione e parole chiave ottimizzate.")

book_info = st.text_area("Di cosa parla il tuo libro? (Titolo, genere, trama breve...)")

if st.button("Genera Contenuti con AI"):
    if not api_key:
        st.error("Inserisci la chiave API di OpenAI nella barra laterale!")
    elif not book_info:
        st.warning("Scrivi qualcosa sul tuo libro prima di generare.")
    else:
        try:
            client = OpenAI(api_key=api_key)
            
            prompt = f"""
            Sei un esperto di marketing per Amazon KDP. 
            Per il seguente libro: '{book_info}'
            1. Scrivi una descrizione accattivante che includa tag HTML (<b>, <i>, <ul>) come richiesto da Amazon.
            2. Fornisci i 7 migliori 'keyword phrases' (parole chiave) per la SEO di Amazon.
            Rispondi in lingua Italiana.
            """
            
            completion = client.chat.completions.create(
                model="gpt-4o", # Modello più avanzato
                messages=[{"role": "user", "content": prompt}]
            )
            
            st.success("Contenuto Generato!")
            st.markdown(completion.choices[0].message.content)
            
        except Exception as e:
            st.error(f"Errore con OpenAI: {e}")

st.divider()

# --- SEZIONE 2: FORMATTAZIONE KDP ---
st.header("2. Formattazione File per Cartaceo")
st.write("Trasforma il tuo documento in formato standard **6x9 pollici** (il più usato su KDP).")

uploaded_file = st.file_uploader("Carica il tuo file Word (.docx)", type="docx")

if uploaded_file:
    if st.button("Formatta Documento"):
        doc = Document(uploaded_file)
        
        # Impostazioni Standard KDP 6x9 (senza bleed)
        for section in doc.sections:
            section.page_width = Inches(6)
            section.page_height = Inches(9)
            
            # Margini di sicurezza per la rilegatura (Gutter)
            section.top_margin = Inches(0.75)
            section.bottom_margin = Inches(0.75)
            section.left_margin = Inches(0.8)  # Margine interno più largo per la colla
            section.right_margin = Inches(0.5) # Margine esterno
        
        # Salvataggio in memoria
        out_buffer = io.BytesIO()
        doc.save(out_buffer)
        out_buffer.seek(0)
        
        st.balloons()
        st.download_button(
            label="Scarica il libro formattato",
            data=out_buffer,
            file_name="libro_kdp_6x9.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
