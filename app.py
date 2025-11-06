import streamlit as st
import pandas as pd
import numpy as np
import time
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ==========================
# ⚙️ CONFIGURAZIONE PAGINA
# ==========================
st.set_page_config(
    page_title="Report Attività Clienti - EdiliziAcrobatica",
    page_icon="fav.png",
    layout="centered"
)

# ==========================
# 🎨 STILE PERSONALIZZATO
# ==========================
st.markdown("""
    <style>
        .block-container { padding-top: 1rem; }
        h1, h2, h3, p { font-family: 'Segoe UI', sans-serif; }
        footer { visibility: hidden; }
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

st.markdown("<div style='height:4px; background-color:#004C97; margin-bottom:25px;'></div>", unsafe_allow_html=True)
st.markdown("<h1 style='text-align:center; color:#000;'>📊 Report Attività Clienti</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center; color:gray; font-size:16px;'>Generatore report automatico – <b>Solo per uso interno EdiliziAcrobatica S.p.A.</b></p>", unsafe_allow_html=True)

oggi = datetime.now().strftime("%d %B %Y")
st.markdown(f"<p style='text-align:center; color:#004C97; font-size:14px;'>🕒 Ultimo aggiornamento: {oggi} – Versione 1.0</p>", unsafe_allow_html=True)

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
    progress_text = "⏳ Elaborazione in corso... attendere."
    my_bar = st.progress(0, text=progress_text)
    start_time = time.time()

    # 1️⃣ Lettura file
    att = pd.read_excel(file_att)
    tab_raw = pd.read_excel(file_tab, header=None, skiprows=3)
    tab_raw.columns = tab_raw.iloc[0]
    tab = tab_raw.drop(0).reset_index(drop=True)
    tab = tab.rename(columns={"macroarea": "Macroarea"})
    my_bar.progress(10, text="📄 File caricati con successo...")

    # 2️⃣ Normalizzazione
    def normalize_name(x):
        if pd.isna(x): return ""
        x = str(x).lower().replace(".", " ").replace("*", " ").replace(",", " ")
        return " ".join(x.split())

    att["NomeSoggetto_n"] = att["NomeSoggetto"].apply(normalize_name)
    tab["Cliente_n"] = tab["Cliente"].apply(normalize_name)

    if "Tipo" in tab.columns:
        def fix_tipo(x):
            x = str(x).strip().capitalize()
            if x.lower().startswith("amministrator"):
                return "Amministratori"
            return x
        tab["Tipo"] = tab["Tipo"].apply(fix_tipo)
    else:
        tab["Tipo"] = "Amministratori"

    my_bar.progress(25, text="🔎 Normalizzazione nomi completata...")

    # 3️⃣ Priorità
    priorita = {
        "04 RICHIESTE": 1, "06 PREVENTIVI": 2, "03 INCONTRI": 3,
        "07 DELIBERE": 4, "05 SOPRALLUOGHI": 5, "01 TELEFONATE": 6, "02 APPUNTAMENTI": 7
    }
    att["Priorita"] = att["Classe Attività"].map(priorita).fillna(999)

    # 4️⃣ Match attività-clienti
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

    my_bar.progress(60, text="📊 Match attività completato...")

    # 5️⃣ Crea DataFrame finale
    database = pd.DataFrame(righe_output).replace({np.nan: ""})

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

    # 6️⃣ Scrivi Excel base
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        database.to_excel(writer, sheet_name="Database", index=False)
        for tipo, grp in sorted(database.groupby("Tipo"), key=lambda x: str(x[0])):
            nome = str(tipo).strip().capitalize() or "Senzatipo"
            grp[
                ["Sede", "Responsabile gestionale", "Cliente", "Anno", "Mese",
                 "Ultima attività", "Da riassegnare",
                 "PREVENTIVATO€", "DELIBERATO€", "FATTURATO€", "INCASSATO€"]
            ].sort_values("Cliente").to_excel(writer, sheet_name=nome, index=False)

    my_bar.progress(75, text="🎨 Applicazione formattazione Excel...")

    # 7️⃣ Formattazione workbook
    wb = load_workbook(output)
    thin = Side(border_style="thin", color="D9D9D9")
    header_fill = PatternFill(start_color="004C97", end_color="004C97", fill_type="solid")
    alt_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
    green_fill = PatternFill(start_color="A6F3A6", end_color="A6F3A6", fill_type="solid")

    for ws in wb.worksheets:
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                if cell.row % 2 == 0:
                    cell.fill = alt_fill
                if cell.value == "Sì":
                    cell.fill = red_fill
                elif cell.value == "No":
                    cell.fill = green_fill
                cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
                cell.alignment = Alignment(horizontal="center", vertical="center")

        for col_cells in ws.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 45)

    if "Amministratori" in wb.sheetnames:
        wb.active = wb.sheetnames.index("Amministratori")

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    my_bar.progress(100, text="✅ File pronto per il download!")

    elapsed = round(time.time() - start_time, 2)
    st.success(f"✅ File elaborato e formattato in {elapsed} secondi!")

    st.download_button(
        label="📥 Scarica il report Excel formattato",
        data=final_output,
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
