import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# ============================================================
# 🧱 CONFIGURAZIONE PAGINA
# ============================================================
st.set_page_config(
    page_title="Report Attività Clienti - EdiliziAcrobatica",
    page_icon="🏗️",
    layout="centered"
)

# ============================================================
# 🎨 STILE GRAFICO
# ============================================================
st.markdown("""
    <style>
    .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        max-width: 900px;
    }

    h1 {
        text-align: center;
        color: #004C97;
        font-weight: 900;
        margin-bottom: 1rem;
    }

    .upload-box {
        background-color: #F8F9FA;
        padding: 20px 25px;
        border: 2px solid #004C97;
        border-radius: 12px;
        margin-top: 1rem;
        margin-bottom: 1.5rem;
    }

    .upload-box h3 {
        color: #004C97;
        margin-bottom: 0.6rem;
    }

    .upload-box p {
        font-size: 15px;
        margin-top: 0;
        margin-bottom: 0.5rem;
        line-height: 1.4;
    }

    div.stButton > button:first-child {
        background-color: #004C97;
        color: white;
        font-weight: bold;
        padding: 0.6rem 2rem;
        border-radius: 10px;
        border: none;
    }
    div.stButton > button:hover {
        background-color: #0062C4;
        color: white;
    }

    div.stDownloadButton > button {
        background-color: #198754 !important;
        color: white !important;
        font-weight: bold;
        border-radius: 10px;
        padding: 0.6rem 1.8rem;
    }
    div.stDownloadButton > button:hover {
        background-color: #157347 !important;
    }

    hr {
        border: 1px solid #004C97;
        margin-top: 1.5rem;
        margin-bottom: 1.5rem;
    }

    footer {
        text-align: center;
        font-size: 13px;
        color: #888;
        margin-top: 2rem;
    }
    </style>
""", unsafe_allow_html=True)


# ============================================================
# ⚙️ FUNZIONE DI ELABORAZIONE (SEMPLIFICATA)
# ============================================================
def genera_report(file_attivita, file_tabella):
    # Legge i due file caricati
    att = pd.read_excel(file_attivita)
    tab = pd.read_excel(file_tabella)

    # Crea un file Excel fittizio per demo
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        att.to_excel(writer, index=False, sheet_name="Attivita")
        tab.to_excel(writer, index=False, sheet_name="Tabella Clienti")
    output.seek(0)
    return output


# ============================================================
# 🖥️ INTERFACCIA UTENTE
# ============================================================
st.image("https://www.ediliziacrobatica.com/wp-content/uploads/2022/05/logo.svg", width=230)
st.markdown("<h1>🏗️ Generatore Report Attività Clienti</h1>", unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# --- SEZIONE ATTIVITÀ ---
st.markdown("""
<div class="upload-box">
<h3>📘 Carica il file delle attività</h3>
<p>
Scaricalo dalla <b>Dashboard Commerciale → Sottoprodotti → Tab Grafici Attività</b><br>
seleziona <b>l'ultimo elenco prima del grafico “Delibere”</b>.<br>
➡️ <b>Attendi il caricamento dei dati e premi “Crea Excel”</b>
</p>
</div>
""", unsafe_allow_html=True)

file_att = st.file_uploader("📘 Seleziona il file delle attività (.xlsx)", type=["xlsx"], help="Esempio: attivita_2025.xlsx")

# --- SEZIONE CLIENTI ---
st.markdown("""
<div class="upload-box">
<h3>📗 Carica il file dei clienti</h3>
<p>
Scaricalo dalla <b>Dashboard Commerciale → Riepilogo Clienti</b>, impostando il periodo <b>dal 2017 ad oggi</b>.<br>
Scarica Excel da <b>“Tabella Clienti (no filtro data)”</b> in fondo alla pagina,<br>
dopo aver atteso il caricamento dei dati.
</p>
</div>
""", unsafe_allow_html=True)

file_tab = st.file_uploader("📗 Seleziona la tabella clienti (.xlsx)", type=["xlsx"], help="Esempio: Tabella_Clienti.xlsx")

# --- BOTTONE ELABORAZIONE ---
if file_att and file_tab:
    if st.button("🚀 Crea file di output"):
        with st.spinner("Elaborazione in corso, attendere..."):
            result = genera_report(file_att, file_tab)
        st.success("✅ File generato con successo!")
        st.download_button(
            label="⬇️ Scarica file Excel formattato",
            data=result,
            file_name=f"output_attivita_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ============================================================
# 🧾 FOOTER
# ============================================================
st.markdown("""
<footer>
© EdiliziAcrobatica – Generatore Report Attività Clienti<br>
Tool interno per la rete commerciale
</footer>
""", unsafe_allow_html=True)
