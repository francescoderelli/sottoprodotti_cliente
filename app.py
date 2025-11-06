import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from difflib import get_close_matches
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime
import time

# ==========================
# ⚙️ CONFIGURAZIONE PAGINA
# ==========================
st.set_page_config(
    page_title="Report Attività Clienti - EdiliziAcrobatica",
    page_icon="fav.png",   # favicon nella tab del browser
    layout="centered"
)

# ==========================
# 🎨 STILE PERSONALIZZATO
# ==========================
st.markdown("""
    <style>
        .block-container {
            padding-top: 1rem;
            padding-bottom: 0rem;
        }
        h1, h2, h3, p {
            font-family: 'Segoe UI', sans-serif;
        }
        footer {
            visibility: hidden;
        }
        .intro {
            background-color: #004C97;
            color: white;
            text-align: center;
            padding: 10px 0px;
            border-radius: 8px;
            font-size: 18px;
            margin-bottom: 25px;
        }
    </style>
""", unsafe_allow_html=True)

# ==========================
# 🏗️ HEADER E BRANDING
# ==========================
col1, col2, col3 = st.columns([1, 3, 1])
with col2:
    st.image("logo.png", width=240)

# Barra blu aziendale
st.markdown("<div style='height:4px; background-color:#004C97; margin-bottom:25px;'></div>", unsafe_allow_html=True)

# Titolo e sottotitolo
st.markdown("<h1 style='text-align: center; color:#000;'>📊 Report Attività Clienti</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center; color:gray; font-size:16px;'>Generatore report automatico – <b>Solo per uso interno EdiliziAcrobatica S.p.A.</b></p>", unsafe_allow_html=True)

oggi = datetime.now().strftime("%d %B %Y")
st.markdown(f"<p style='text-align:center; color:#004C97; font-size:14px; margin-top:-10px;'>🕒 Ultimo aggiornamento: {oggi} – Versione 1.0</p>", unsafe_allow_html=True)

# ==========================
# 📘 INTRO COLORATA
# ==========================
st.markdown("<div class='intro'>Benvenuto nel generatore report attività clienti</div>", unsafe_allow_html=True)

# ==========================
# 📂 UPLOAD FILE
# ==========================
st.markdown("### 📁 Carica i tuoi file Excel")

st.markdown("#### 📄 File Attività")
st.markdown("""
Scaricalo dalla **Dashboard Commerciale → Sottoprodotti → Tab Grafici Attività**,  
seleziona l’ultimo elenco prima del grafico *“Delibere”*.  
➡️ Attendi il caricamento dei dati e premi **Crea Excel**.
""")

file_att = st.file_uploader("Seleziona il file delle attività (.xlsx)", type=["xlsx"], key="att")

st.markdown("---")
st.markdown("#### 📗 File Clienti")
st.markdown("""
Scaricalo dalla **Dashboard Commerciale → Riepilogo Clienti**,  
impostando il periodo **dal 2017 ad oggi**,  
e scarica Excel da *“Tabella Clienti (no filtro data)”* in fondo alla pagina,  
dopo aver atteso il caricamento dei dati.
""")

file_tab = st.file_uploader("Seleziona la tabella clienti (.xlsx)", type=["xlsx"], key="cli")

# ==========================
# 🚀 ELABORAZIONE FILE
# ==========================
if file_att and file_tab:
    st.markdown("---")
    st.info("⏳ Elaborazione in corso... attendere qualche istante.")
    start_time = time.time()

    att = pd.read_excel(file_att)
    tab_raw = pd.read_excel(file_tab, header=None, skiprows=3)
    tab_raw.columns = tab_raw.iloc[0]
    tab = tab_raw.drop(0).reset_index(drop=True)

    tab = tab.rename(columns={"macroarea": "Macroarea"})

    # Normalizza i nomi
    def normalize_name(x):
        if pd.isna(x): return ""
        x = str(x).lower().replace(".", " ").replace("*", " ").replace(",", " ")
        return " ".join(x.split())

    att["NomeSoggetto_n"] = att["NomeSoggetto"].apply(normalize_name)
    tab["Cliente_n"] = tab["Cliente"].apply(normalize_name)

    # Tipo
    if "Tipo" in tab.columns:
        def fix_tipo(x):
            x = str(x).strip().capitalize()
            if x.lower().startswith("amministrator"):
                return "Amministratori"
            return x
        tab["Tipo"] = tab["Tipo"].apply(fix_tipo)
    else:
        tab["Tipo"] = "Amministratori"

    # Priorità
    priorita = {
        "04 RICHIESTE": 1,
        "06 PREVENTIVI": 2,
        "03 INCONTRI": 3,
        "07 DELIBERE": 4,
        "05 SOPRALLUOGHI": 5,
        "01 TELEFONATE": 6,
        "02 APPUNTAMENTI": 7
    }
    att["Priorita"] = att["Classe Attività"].map(priorita).fillna(999)

    righe_output = []
    for _, r in tab.iterrows():
        cliente_norm = r["Cliente_n"]
        tipo_cli = r["Tipo"]
        sede_cli = r.get("Sede", "")
        resp_gest = r.get("Responsabile", "")

        att_cli = att[att["NomeSoggetto_n"] == cliente_norm]
        if att_cli.empty and cliente_norm:
            invertito = " ".join(cliente_norm.split()[::-1])
            att_cli = att[att["NomeSoggetto_n"] == invertito]

        if not att_cli.empty:
            att_cli = att_cli.sort_values(["Anno", "Mese", "Priorita"]).iloc[-1]
            anno_att, mese_att = int(att_cli["Anno"]), int(att_cli["Mese"])
            diff_mesi = (2025 - anno_att) * 12 + (11 - mese_att)
            da_ria = "Sì" if diff_mesi > 2 else "No"
            righe_output.append({
                "Sede": sede_cli,
                "Responsabile gestionale": resp_gest,
                "Cliente": r["Cliente"],
                "Anno": anno_att,
                "Mese": mese_att,
                "Ultima attività": att_cli["Classe Attività"],
                "Da riassegnare": da_ria,
                "PREVENTIVATO€": r.get("PREVENTIVATO€", ""),
                "DELIBERATO€": r.get("DELIBERATO€", ""),
                "FATTURATO€": r.get("FATTURATO€", ""),
                "INCASSATO€": r.get("INCASSATO€", ""),
                "Tipo": tipo_cli
            })
        else:
            righe_output.append({
                "Sede": sede_cli,
                "Responsabile gestionale": resp_gest,
                "Cliente": r["Cliente"],
                "Anno": "",
                "Mese": "",
                "Ultima attività": "",
                "Da riassegnare": "Sì",
                "PREVENTIVATO€": r.get("PREVENTIVATO€", ""),
                "DELIBERATO€": r.get("DELIBERATO€", ""),
                "FATTURATO€": r.get("FATTURATO€", ""),
                "INCASSATO€": r.get("INCASSATO€", ""),
                "Tipo": tipo_cli
            })

    database = pd.DataFrame(righe_output).replace({np.nan: ""})

    # Formatting
    def format_euro(x):
        if x == "" or pd.isna(x): return ""
        try:
            val = float(str(x).replace("€", "").replace(".", "").replace(",", "."))
            return f"€ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return str(x)

    for c in ["PREVENTIVATO€", "DELIBERATO€", "FATTURATO€", "INCASSATO€"]:
        if c in database.columns:
            database[c] = database[c].apply(format_euro)

    # Salva
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        database.to_excel(writer, sheet_name="Database", index=False)
        for tipo, grp in sorted(database.groupby("Tipo"), key=lambda x: str(x[0])):
            nome = str(tipo).strip().capitalize() or "Senzatipo"
            grp[["Sede", "Responsabile gestionale", "Cliente", "Anno", "Mese",
                 "Ultima attività", "Da riassegnare",
                 "PREVENTIVATO€", "DELIBERATO€", "FATTURATO€", "INCASSATO€"]
                ].sort_values("Cliente").to_excel(writer, sheet_name=nome, index=False)

    # Timer
    elapsed = round(time.time() - start_time, 2)
    st.success(f"✅ File elaborato correttamente in {elapsed} secondi!")

    st.download_button(
        label="📥 Scarica il report Excel",
        data=output.getvalue(),
        file_name="report_attivita_clienti.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ==========================
# 📜 FOOTER
# ==========================
st.markdown("---")
st.markdown("""
<p style='text-align:center; color:gray; font-size:13px;'>
© 2025 <b>EdiliziAcrobatica S.p.A.</b> – Tutti i diritti riservati.<br>
Uso interno, vietata la diffusione esterna.
</p>
""", unsafe_allow_html=True)
