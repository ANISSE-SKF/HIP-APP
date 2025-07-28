import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.styles import Alignment

st.set_page_config(page_title="PDF → Excel Extractor", layout="centered")
st.title("📄 PDF → Excel Extracteur de Données")
st.markdown("Dépose un fichier PDF pour extraire les données et générer un fichier Excel structuré.")

def extract_values_from_text(text):
    data = {
        "BAR": [], "DIAMETER": [], "Elong.4D": [], "Elong.5D": [], "InitialD": [],
        "Proof(0.2%)": [], "mE": [], "RT UTS": [], "450°C UTS": [], "RT 0.2%Proof": [],
        "450°C 0.2%Proof": [], "ElongatFracture": [], "ElongafterFracture": [],
        "HRC": [], "Moyenne_HRC": []
    }

    # BAR and DIAMETER
    match = re.search(r"CAST\*.*?([A-Z0-9]+)\s+Serial No\. ([0-9/]+)", text)
    if match:
        data["BAR"] = [match.group(1), ""]
        data["DIAMETER"] = [match.group(2), ""]

    # RT UTS
    data["RT UTS"] = re.findall(r"RT.*?UTS.*?≥ \d+\n(\d+)", text)[:2]
    # 450°C UTS
    data["450°C UTS"] = re.findall(r"450°C.*?UTS.*?≥ \d+\n([\d.]+)", text)[:2]
    # RT 0.2%Proof
    data["RT 0.2%Proof"] = re.findall(r"RT.*?0\.2% Proof.*?≥ \d+\n(\d+)", text)[:2]
    # 450°C 0.2%Proof
    data["450°C 0.2%Proof"] = re.findall(r"450°C.*?0\.2% Proof.*?≥ \d+\n([\d.]+)", text)[:2]
    # ElongatFracture
    data["ElongatFracture"] = re.findall(r"RT.*?Elong at Fracture.*?(\d+)%", text)[:2]
    # ElongafterFracture
    data["ElongafterFracture"] = re.findall(r"450°C.*?Elong after Fracture.*?(\d+\.?\d*)%", text)[:2]
    # HRC
    hrc_values = re.findall(r"HRC.*?\n(\d+)\n(\d+)\n(\d+)", text)
    if hrc_values:
        hrc = hrc_values[0]
        data["HRC"] = [", ".join(hrc), ""]
        moyenne = round(sum(map(int, hrc)) / 3)
        data["Moyenne_HRC"] = [str(moyenne), ""]

    # Fill missing values with empty strings
    for key in data:
        while len(data[key]) < 2:
            data[key].append("")

    return data

def create_excel(data_dict):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active

    # Fusion des en-têtes
    ws.merge_cells("A1:G1")
    ws["A1"] = "Curve"
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("H1:O1")
    ws["H1"] = "Special Test"
    ws["H1"].alignment = Alignment(horizontal="center", vertical="center")

    headers = list(data_dict.keys())
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_num)
        cell.value = header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx in range(2):
        for col_idx, key in enumerate(headers, 1):
            ws.cell(row=row_idx + 3, column=col_idx).value = data_dict[key][row_idx]

    wb.save(output)
    output.seek(0)
    return output

uploaded_file = st.file_uploader("Dépose un fichier PDF", type="pdf")

if uploaded_file:
    with st.spinner("📤 Traitement du fichier..."):
        pdf_bytes = uploaded_file.read()
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            full_text = "\n".join([page.get_text() for page in doc])

        extracted_data = extract_values_from_text(full_text)
        excel_file = create_excel(extracted_data)

        st.success("✅ Fichier Excel généré avec succès !")
        st.download_button(
            label="📥 Télécharger le fichier Excel",
            data=excel_file,
            file_name="resultats.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


