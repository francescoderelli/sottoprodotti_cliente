import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ============================================================
# CONFIGURAZIONE PAGINA
# ============================================================
st.set_page_config(
    page_title="Report Attività Clienti - EdiliziAcrobatica",
    page_icon="🏗️",
    layout="centered"
)

# ============================================================
# STILE GRAFICO
# ============================================================
st.markdown("""
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 900px;}
    h1 {text-align:center; color:#004C97; font-weight:900; margin-bottom:1rem;}
    .upload-box {
        background-color:#F8F9FA; padding:20px 25px;
        border:2px solid #004C97; border-radius:12px;
        margin-top:1rem; margin-bottom:1.5rem;
    }
    .upload-box h3 {color:#004C97; margin-bottom:0.6rem;}
    .upload-box p {font-size:15px; margin-top:0; margin-bottom:0.5rem; line-height:1.4;}
    div.stButton > button:first-child {
        background-color:#004C97; color:white; font-weight:bold;
        padding:0.6rem 2rem; border-radius:10px; border:none;
    }
    div.stButton > button:hover {background-color:#0062C4; color:white;}
    div.stDownloadButton > button {
        background-color:#198754 !important; color:white !important;
        font-weight:bold; border-radius:10px; padding:0.6rem 1.8rem;
    }
    div.stDownloadButton > button:hover {background-color:#157347 !important;}
    hr {border:1px solid #004C97; margin-top:1.5rem; margin-bottom:1.5rem;}
    footer {text-align:center; font-size:13px; color:#888; margin-top:2rem;}
    </style>
""", unsafe_allow_html=True)


# ============================================================
# FUNZIONI DI CONTROLLO FILE
# ============================================================

def check_file_attivita(df):
    """Verifica che il file attività contenga le colonne corrette"""
    colonne_richieste = ['Anno', 'Mese', 'Classe Attività', 'Responsabile', 'Sede', 'NomeSoggetto']
    colonne_presenti = [c.strip() for c in df.columns]
    mancanti = [c for c in colonne_richieste if c not in colonne_presenti]
    if mancanti:
        return False, f"❌ File Attività non valido. Mancano le colonne: {', '.join(mancanti)}"
    return True, "✅ File Attività corretto."


def check_file_clienti(df):
    """Verifica che il file clienti contenga le colonne corrette"""
    colonne_richieste = ['Country', 'Rete', 'Macroarea', 'Regione', 'Provincia', 'Sede', 'Responsabile', 'Cliente']
    colonne_presenti = [c.strip() for c in df.columns]
    mancanti = [c for c in colonne_richieste if c not in colonne_presenti]
    if mancanti:
        return False, f"❌ File Clienti non valido. Mancano le colonne: {', '.join(mancanti)}"
    return True, "✅ File Clienti corretto."


# ============================================================
# FUNZIONE DI ELABORAZIONE
# ============================================================

def genera_report(file_attivita, file_clienti):

    # --- Lettura file attività ---
    attivita = pd.read_excel(file_attivita)
    ok_a, msg_a = check_file_attivita(attivita)
    if not ok_a:
        st.error(msg_a)
        return None
    st.success(msg_a)

    # --- Lettura file clienti (salta 3 righe) ---
    try:
        tabella = pd.read_excel(file_clienti, header=None, skiprows=3)
        tabella.columns = tabella.iloc[0]
        tabella = tabella.drop(0).reset_index(drop=True)
    except Exception as e:
        st.error("❌ Errore nella lettura del file clienti. Verifica che sia scaricato da 'Tabella Clienti (no filtro data)'.")
        return None

    ok_c, msg_c = check_file_clienti(tabella)
    if not ok_c:
        st.error(msg_c)
        return None
    st.success(msg_c)

    # --- Pulizia ---
    attivita = attivita.replace({np.nan: ""})
    tabella = tabella.replace({np.nan: ""})

    # --- Normalizza nomi per match ---
    def normalize_name(x):
        if pd.isna(x): return ""
        x = str(x).lower().replace(".", " ").replace("*", " ").replace(",", " ")
        return " ".join(x.split())

    attivita["NomeSoggetto_n"] = attivita["NomeSoggetto"].apply(normalize_name)
    tabella["Cliente_n"] = tabella["Cliente"].apply(normalize_name)

    # --- Tipo (colonna P) ---
    if "Tipo" in tabella.columns:
        def fix_tipo(x):
            x = str(x).strip().capitalize()
            if x.lower().startswith("amministrator"):
                return "Amministratori"
            return x
        tabella["Tipo"] = tabella["Tipo"].apply(fix_tipo)
    else:
        tabella["Tipo"] = "Amministratori"

    # --- Priorità attività ---
    priorita = {
        "04 RICHIESTE": 1, "06 PREVENTIVI": 2, "03 INCONTRI": 3,
        "07 DELIBERE": 4, "05 SOPRALLUOGHI": 5, "01 TELEFONATE": 6, "02 APPUNTAMENTI": 7
    }
    attivita["Priorita"] = attivita["Classe Attività"].map(priorita).fillna(999)

    # --- Matching attività/clienti ---
    match_rows = []
    progress = st.progress(0)
    for i, (_, riga_cli) in enumerate(tabella.iterrows()):
        progress.progress(int((i + 1) / len(tabella) * 100))

        cliente_norm = riga_cli["Cliente_n"]
        tipo_cli = riga_cli["Tipo"]
        resp_gest = riga_cli.get("Responsabile", "")
        sede_cli = riga_cli.get("Sede", "")

        att_match = attivita[attivita["NomeSoggetto_n"] == cliente_norm]
        if att_match.empty and cliente_norm:
            invertito = " ".join(cliente_norm.split()[::-1])
            att_match = attivita[attivita["NomeSoggetto_n"] == invertito]

        if not att_match.empty:
            att_match = att_match.sort_values(by=["Anno", "Mese", "Priorita"]).iloc[-1]
            anno_att, mese_att = int(att_match["Anno"]), int(att_match["Mese"])
            diff_mesi = (2025 - anno_att) * 12 + (11 - mese_att)
            da_riass = "Sì" if diff_mesi > 2 else "No"

            match_rows.append({
                "Sede": sede_cli,
                "Responsabile gestionale": resp_gest,
                "Cliente": riga_cli["Cliente"],
                "Anno": anno_att,
                "Mese": mese_att,
                "Ultima attività": att_match["Classe Attività"],
                "Da riassegnare": da_riass,
                "PREVENTIVATO€": riga_cli.get("PREVENTIVATO€", ""),
                "DELIBERATO€": riga_cli.get("DELIBERATO€", ""),
                "FATTURATO€": riga_cli.get("FATTURATO€", ""),
                "INCASSATO€": riga_cli.get("INCASSATO€", ""),
                "Tipo": tipo_cli
            })
        else:
            match_rows.append({
                "Sede": sede_cli,
                "Responsabile gestionale": resp_gest,
                "Cliente": riga_cli["Cliente"],
                "Anno": "",
                "Mese": "",
                "Ultima attività": "",
                "Da riassegnare": "Sì",
                "PREVENTIVATO€": riga_cli.get("PREVENTIVATO€", ""),
                "DELIBERATO€": riga_cli.get("DELIBERATO€", ""),
                "FATTURATO€": riga_cli.get("FATTURATO€", ""),
                "INCASSATO€": riga_cli.get("INCASSATO€", ""),
                "Tipo": tipo_cli
            })

    database = pd.DataFrame(match_rows)
    database = database.replace({np.nan: ""})

    # --- Formattazione numeri € ---
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

    # --- Esporta in Excel ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        database.to_excel(writer, sheet_name="Database", index=False)
        for tipo, grp in sorted(database.groupby("Tipo"), key=lambda x: str(x[0])):
            nome = str(tipo).strip().capitalize() or "SenzaTipo"
            grp.to_excel(writer, sheet_name=nome, index=False)

    # --- Formattazione estetica ---
    output.seek(0)
    wb = load_workbook(output)
    thin = Side(border_style="thin", color="D9D9D9")
    header_fill = PatternFill(start_color="004C97", end_color="004C97", fill_type="solid")
    alt_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    for ws in wb.worksheets:
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                if cell.row % 2 == 0: cell.fill = alt_fill
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
    return final_output


# ============================================================
# INTERFACCIA UTENTE
# ============================================================
st.image("https://www.ediliziacrobatica.com/wp-content/uploads/2022/05/logo.svg", width=230)
st.markdown("<h1>🏗️ Generatore Report Attività Clienti</h1>", unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

st.markdown("""
<div class="upload-box">
<h3>📘 Carica il file delle attività</h3>
<p>Scaricalo dalla <b>Dashboard Commerciale → Sottoprodotti → Tab Grafici Attività</b><br>
seleziona <b>l'ultimo elenco prima del grafico “Delibere”</b>.<br>
➡️ <b>Attendi il caricamento dei dati e premi “Crea Excel”</b></p>
</div>
""", unsafe_allow_html=True)
file_att = st.file_uploader("📘 Seleziona il file delle attività (.xlsx)", type=["xlsx"])

st.markdown("""
<div class="upload-box">
<h3>📗 Carica il file dei clienti</h3>
<p>Scaricalo dalla <b>Dashboard Commerciale → Riepilogo Clienti</b>, periodo <b>dal 2017 ad oggi</b>.<br>
Scarica Excel da <b>“Tabella Clienti (no filtro data)”</b> dopo aver atteso il caricamento dei dati.</p>
</div>
""", unsafe_allow_html=True)
file_tab = st.file_uploader("📗 Seleziona la tabella clienti (.xlsx)", type=["xlsx"])

if file_att and file_tab:
    if st.button("🚀 Crea file di output"):
        with st.spinner("Elaborazione in corso, attendere..."):
            result = genera_report(file_att, file_tab)
        if result:
            st.success("✅ File generato con successo!")
            st.download_button(
                "⬇️ Scarica file Excel formattato",
                data=result,
                file_name=f"output_attivita_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.markdown("""
<footer>© EdiliziAcrobatica – Generatore Report Attività Clienti<br>
Tool interno per la rete commerciale</footer>
""", unsafe_allow_html=True)
