import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from difflib import get_close_matches
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import time

# ==========================
# 0️⃣ Configurazione pagina
# ==========================
st.set_page_config(
    page_title="Report Attività Clienti - EdiliziAcrobatica",
    page_icon="📊",
    layout="centered"
)

# ==========================
# 1️⃣ Header e Logo
# ==========================
st.image(
    "https://www.ediliziacrobatica.com/wp-content/uploads/2021/02/logo-acrobatica.svg",
    width=250
)
st.title("📊 Report Attività Clienti")
st.caption("Generatore report automatico – Solo per uso interno EdiliziAcrobatica S.p.A.")

# ==========================
# 2️⃣ Istruzioni a video
# ==========================
st.markdown("""
### 📘 File Attività
Scaricalo dalla **Dashboard Commerciale → Sottoprodotti → Tab Grafici Attività**,  
seleziona **l’ultimo elenco prima del grafico “Delibere”**.  
➡️ Attendi il caricamento dei dati e premi **Crea Excel**.
""")

file_att = st.file_uploader("📂 Seleziona il file delle attività (.xlsx)", type=["xlsx"])

st.markdown("""
### 📗 File Clienti
Scaricalo dalla **Dashboard Commerciale → Riepilogo Clienti**,  
impostando il periodo **dal 2017 ad oggi**,  
e scarica Excel da **“Tabella Clienti (no filtro data)”** in fondo alla pagina,  
dopo aver atteso il caricamento dei dati.
""")

file_tab = st.file_uploader("📂 Seleziona la tabella clienti (.xlsx)", type=["xlsx"])

# ==========================
# 3️⃣ Funzioni utili
# ==========================
def normalize_name(x):
    if pd.isna(x): return ""
    x = str(x).lower().replace(".", " ").replace("*", " ").replace(",", " ")
    return " ".join(x.split())

def to_float_euro(x):
    if pd.isna(x) or str(x).strip() == "":
        return np.nan
    x = str(x).replace("€", "").replace(" ", "")
    if "," in x and "." in x:
        x = x.replace(".", "").replace(",", ".")
    elif "," in x:
        x = x.replace(",", ".")
    try:
        return float(x)
    except ValueError:
        return np.nan

def format_euro(x):
    if pd.isna(x) or x == "":
        return ""
    try:
        return f"€ {float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(x)

# ==========================
# 4️⃣ Elaborazione file
# ==========================
if file_att and file_tab:
    start_time = time.time()
    st.info("⏳ Elaborazione in corso... Attendere qualche secondo...")

    # Lettura file
    att = pd.read_excel(file_att)
    tab_raw = pd.read_excel(file_tab, header=None, skiprows=3)
    tab_raw.columns = tab_raw.iloc[0]
    tab = tab_raw.drop(0).reset_index(drop=True)
    tab = tab.rename(columns={"macroarea": "Macroarea"})

    # Normalizzazione nomi
    att["NomeSoggetto_n"] = att["NomeSoggetto"].apply(normalize_name)
    tab["Cliente_n"] = tab["Cliente"].apply(normalize_name)

    # Tipo cliente
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
    tutti_nomi_att = att["NomeSoggetto_n"].unique().tolist()
    totale = len(tab)
    progress_text = st.empty()
    progress_bar = st.progress(0)

    # ==========================
    #  Match intelligente
    # ==========================
    for i, r in tab.iterrows():
        cliente_norm = r["Cliente_n"]
        att_cli = att[att["NomeSoggetto_n"] == cliente_norm]

        # Prova nome invertito
        if att_cli.empty and cliente_norm:
            invertito = " ".join(cliente_norm.split()[::-1])
            att_cli = att[att["NomeSoggetto_n"] == invertito]

        # Prova similitudine fuzzy (con soglia bassa per “Facincani” e “A.CON”)
        if att_cli.empty and cliente_norm:
            simili = get_close_matches(cliente_norm, tutti_nomi_att, n=1, cutoff=0.65)
            if simili:
                att_cli = att[att["NomeSoggetto_n"] == simili[0]]

        # Crea record
        if not att_cli.empty:
            att_cli = att_cli.sort_values(["Anno", "Mese", "Priorita"]).iloc[-1]
            anno_att, mese_att = int(att_cli["Anno"]), int(att_cli["Mese"])
            diff_mesi = (2025 - anno_att) * 12 + (11 - mese_att)
            da_ria = "Sì" if diff_mesi > 2 else "No"
            righe_output.append({
                "Sede": r.get("Sede", ""),
                "Responsabile gestionale": r.get("Responsabile", ""),
                "Cliente": r.get("Cliente", ""),
                "Anno": anno_att,
                "Mese": mese_att,
                "Ultima attività": att_cli["Classe Attività"],
                "Da riassegnare": da_ria,
                "PREVENTIVATO€": r.get("PREVENTIVATO€", ""),
                "DELIBERATO€": r.get("DELIBERATO€", ""),
                "FATTURATO€": r.get("FATTURATO€", ""),
                "INCASSATO€": r.get("INCASSATO€", ""),
                "Tipo": r.get("Tipo", "")
            })
        else:
            righe_output.append({
                "Sede": r.get("Sede", ""),
                "Responsabile gestionale": r.get("Responsabile", ""),
                "Cliente": r.get("Cliente", ""),
                "Anno": "",
                "Mese": "",
                "Ultima attività": "",
                "Da riassegnare": "Sì",
                "PREVENTIVATO€": r.get("PREVENTIVATO€", ""),
                "DELIBERATO€": r.get("DELIBERATO€", ""),
                "FATTURATO€": r.get("FATTURATO€", ""),
                "INCASSATO€": r.get("INCASSATO€", ""),
                "Tipo": r.get("Tipo", "")
            })

        if i % 10 == 0 or i == totale - 1:
            progress_bar.progress((i + 1) / totale)
            progress_text.text(f"🔎 Elaborazione clienti... ({i + 1}/{totale})")

    # Attività senza match (Amministratori da assegnare)
    clienti_norm = set(tab["Cliente_n"].dropna().unique())
    att_no_match = att[~att["NomeSoggetto_n"].isin(clienti_norm)].copy()
    if not att_no_match.empty:
        att_no_match = (
            att_no_match.sort_values(["Anno", "Mese", "Priorita"])
            .groupby("NomeSoggetto", as_index=False)
            .last()
        )
        att_no_match["Da riassegnare"] = "Sì"
        att_no_match["Responsabile gestionale"] = att_no_match["Responsabile"]
        att_no_match["Cliente"] = att_no_match["NomeSoggetto"]
        att_no_match["Ultima attività"] = att_no_match["Classe Attività"]
        att_no_match["Tipo"] = "Amministratori"
        for c in ["PREVENTIVATO€","DELIBERATO€","FATTURATO€","INCASSATO€"]:
            att_no_match[c] = ""
        righe_output.extend(att_no_match[[
            "Sede","Responsabile gestionale","Cliente","Anno","Mese","Ultima attività",
            "Da riassegnare","PREVENTIVATO€","DELIBERATO€","FATTURATO€","INCASSATO€","Tipo"
        ]].to_dict(orient="records"))

    progress_bar.empty()
    progress_text.empty()

    # ==========================
    # DataFrame finale
    # ==========================
    database = pd.DataFrame(righe_output).replace({np.nan: ""})
    for c in ["PREVENTIVATO€","DELIBERATO€","FATTURATO€","INCASSATO€"]:
        database[c] = database[c].apply(to_float_euro).apply(format_euro)

    st.success("✅ Analisi completata! Preparazione file Excel...")

    # ==========================
    # Esporta Excel formattato
    # ==========================
    output = BytesIO()
    col_order = [
        "Sede","Responsabile gestionale","Cliente","Anno","Mese",
        "Ultima attività","Da riassegnare",
        "PREVENTIVATO€","DELIBERATO€","FATTURATO€","INCASSATO€"
    ]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        database.to_excel(writer, sheet_name="Database", index=False)
        for tipo, grp in sorted(database.groupby("Tipo"), key=lambda x: str(x[0])):
            nome = str(tipo).strip().capitalize() or "Senzatipo"
            grp[col_order].sort_values("Cliente").to_excel(writer, sheet_name=nome, index=False)

    # ==========================
    # Formattazione Excel
    # ==========================
    output.seek(0)
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
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                if cell.row % 2 == 0:
                    cell.fill = alt_fill
                if cell.value == "Sì":
                    cell.fill = red_fill
                elif cell.value == "No":
                    cell.fill = green_fill
                cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for col_cells in ws.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 45)

    if "Amministratori" in wb.sheetnames:
        wb.active = wb.sheetnames.index("Amministratori")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    elapsed = time.time() - start_time
    minuti = int(elapsed // 60)
    secondi = int(elapsed % 60)

    st.success(f"🎉 Report completato in {minuti} min {secondi} sec!")
    st.download_button(
        label="📥 Scarica report_attivita_clienti.xlsx",
        data=buf,
        file_name="report_attivita_clienti.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ==========================
# 5️⃣ Footer
# ==========================
st.markdown("""
---
© 2025 **EdiliziAcrobatica S.p.A.**  
Tutti i diritti riservati • Uso interno vietato alla diffusione esterna.
""")
