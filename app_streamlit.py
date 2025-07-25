import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="Extraction PDF vers Excel", layout="centered")
st.title("📄 Convertisseur PDF → Excel")

st.markdown("Dépose ton fichier PDF ci-dessous. Le texte sera extrait et converti en Excel.")

uploaded_file = st.file_uploader("Choisis un fichier PDF", type="pdf")

if uploaded_file:
    with st.spinner("📤 Traitement du fichier..."):
        # Extraire le texte du PDF
        text_data = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    for line in lines:
                        text_data.append([line])

        # Créer un DataFrame
        df = pd.DataFrame(text_data, columns=["Texte extrait"])

        # Sauvegarder dans un fichier Excel en mémoire
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Extrait PDF")
        output.seek(0)

        st.success("✅ Fichier traité avec succès !")
        st.download_button(
            label="📥 Télécharger le fichier Excel",
            data=output,
            file_name="resultats 2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
