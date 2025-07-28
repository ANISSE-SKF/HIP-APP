import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import io

st.set_page_config(page_title="Extraction PDF vers Excel", layout="centered")
st.title("📄 Convertisseur PDF → Excel")

st.markdown("Dépose ton fichier PDF ci-dessous. Les tableaux seront extraits et convertis en Excel.")

uploaded_file = st.file_uploader("Choisis un fichier PDF", type="pdf")

if uploaded_file:
    with st.spinner("📤 Traitement du fichier..."):
        text_data = []

        # Lire le contenu du fichier une seule fois
        pdf_bytes = uploaded_file.read()

        # Ouvrir le PDF avec PyMuPDF
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            for page in doc:
                blocks = page.get_text("blocks")
                for block in blocks:
                    text = block[4].strip()
                    if text:
                        lines = text.split("\n")
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
            file_name="resultats.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
