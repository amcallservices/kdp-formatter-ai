import streamlit as st
from docx import Document
from docx.shared import Inches
import io
import google.generativeai as genai

st.set_page_config(page_title="KDP Toolkit AI", page_icon="📖")

st.title("📖 KDP Formatter & AI Assistant")

# --- PARTE 1: AI PER METADATI ---
st.header("1. Generatore Metadati AI")
api_key = st.sidebar.text_input("Inserisci Google API Key:", type="password")

topic = st.text_input("Di cosa parla il tuo libro?")
if st.button("Genera Descrizione e Tag"):
    if not api_key:
        st.error("Inserisci la chiave API nella barra laterale!")
    else:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-pro')
        prompt = f"Crea una descrizione marketing per Amazon KDP e 7 parole chiave SEO per un libro che parla di: {topic}"
        response = model.generate_content(prompt)
        st.info(response.text)

st.divider()

# --- PARTE 2: FORMATTAZIONE WORD ---
st.header("2. Formattatore Layout (6x9 pollici)")
uploaded_file = st.file_uploader("Carica il tuo file Word (.docx)", type="docx")

if uploaded_file:
    if st.button("Formatta per KDP"):
        doc = Document(uploaded_file)
        
        # Applichiamo i margini standard KDP per un libro senza bleed
        for section in doc.sections:
            section.page_width = Inches(6)
            section.page_height = Inches(9)
            section.top_margin = Inches(0.75)
            section.bottom_margin = Inches(0.75)
            section.left_margin = Inches(0.75) # Margine interno
            section.right_margin = Inches(0.5) # Margine esterno
        
        # Salvataggio
        out_buffer = io.BytesIO()
        doc.save(out_buffer)
        out_buffer.seek(0)
        
        st.success("File pronto per il download!")
        st.download_button(
            label="Scarica Word Formattato",
            data=out_buffer,
            file_name="libro_kdp_6x9.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
