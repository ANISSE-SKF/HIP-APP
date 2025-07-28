import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.styles import Alignment

st.set_page_config(page_title="Convertisseur PDF → Excel", layout="centered")
st.title("📄 Convertisseur PDF → Excel (Format Résultat 4)")
st.markdown("Dépose ton fichier PDF ci-dessous. Les données seront extraites et converties en Excel avec le format exact de `resultats 4.xlsx`.")

def extract_data_from_pdf(pdf_bytes):
    data = {
        "BAR": [], "DIAMETER": [], "Elong.4D": [], "Elong.5D": [], "InitialD": [],
        "Proof(0.2%)": [], "mE": [], "RT UTS": [], "450°C UTS": [], "RT 0.2%Proof": [],
        "450°C 0.2%Proof": [], "ElongatFracture": [], "ElongafterFracture": [],
        "HRC": [], "Moyenne_HRC": []
    }

    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        full_text = "\n".join([page.get_text() for page in doc])

    # Extraction des valeurs
    bar_match = re.search(r"CAST\*[\s\n]+([A-Z0-9]+)[\s\n]+Serial No\. ([0-9/]+)", full_text)
    if bar_match:
        data["BAR"] = [bar_match.group(1), ""]
        data["DIAMETER"] = [bar_match.group(2), ""]

    data["RT UTS"] = re.findall(r"RT.*?UTS.*?≥ \d+\n(\d+)", full_text)[:2]
    data["450°C UTS"] = re.findall(r"450°C.*?UTS.*?≥ \d+\n([\d.]+)", full_text)[:2]
    data["RT 0.2%Proof"] = re.findall(r"RT.*?0\.2% Proof.*?≥ \d+\n(\d+)", full_text)[:2]
    data["450°C 0.2%Proof"] = re.findall(r
