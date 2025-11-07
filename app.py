# ==================================================
# 📊 Report Attività Clienti – EdiliziAcrobatica
# Versione 1.1 • Corpo logico 1.0 invariato
# + Controllo file + Barra avanzamento
# ==================================================

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import time

# ======================
# 1️⃣ CONFIGURAZIONE PAGINA
# ======================
st.set_page_config(
    page_title="Report Attività Clienti - EdiliziAcrobatica",
    page_icon="fav.png",
    layout="centered"
)

# ======================
# 2️⃣ HEADER E LOGO
# ======================
col1, col2, col3 = st.columns([1,3,1])
with col2:
    st.image("logo.png", width=500)

st.markdown(
    "<h2 style='text-align:center; color:#004C97;'>Report Attività Clienti</h2>",
    unsafe_allow_html=True
)
st.caption("Sistema interno di analisi clienti – EdiliziAcrobatica S.p.A.")

# ======================
# 3️⃣ ISTRUZIONI E UPLOAD FILE
# ======================
st.markdown("""
---
### 📘 File Attività
**Scaricalo dalla Dashboard Commerciale → Sottoprodotti → Tab Grafici Attività**  
➡️ Seleziona **l’ultimo elenco prima del grafico “Delibere”**  
🕒 Attendi il caricamento completo dei dati e premi **Crea Excel**
""")
file_att = st.file_uploader("📂 Seleziona il file delle attività (.xlsx)", type=["xlsx"])

st.markdown("""
### 📗 File Clienti
**Scaricalo dalla Dashboard Commerciale → Riepilogo Clienti**  
🗓️ Imposta il periodo **dal 2017 ad oggi**  
⬇️ Scarica Excel da **“Tabella Clienti (no filtro data)”** in fondo alla pagina  
📄 Dopo aver atteso il caricamento dei dati.
""")
file_tab = st.file_uploader("📂 Seleziona la tabella clienti (.xlsx)", type=["xlsx"])

# ======================
# 4️⃣ CONTROLLO STRUTTURA FILE
# ======================
if file_att and file_tab:
    st.info("🧩 Verifica struttura file...")

    try:
        att_check = pd.read_excel(file_att, nrows=5)
        tab_check = pd.read_excel(file_tab, header=None, skiprows=3, nrows=5)

        att_cols = ["Anno","Mese","Classe Attività","Responsabile","NomeSoggetto"]
        tab_valid = tab_check.iloc[4].astype(str).str.contains("Cliente", case=False, na=False).any()

        if all(col in att_check.columns for col in att_cols) and tab_valid:
            st.success("✅ Struttura file corretta!")
        else:
            st.error("❌ Struttura file non conforme. Carica i file originali scaricati dalla Dashboard Commerciale.")
            st.stop()
    except Exception as e:
        st.error(f"Errore durante la verifica file: {e}")
        st.stop()

    # ======================
    # 5️⃣ AVVIO ELABORAZIONE (CORPO 1.0 INVARIATO)
    # ======================
    start_time = time.time()
    progress_bar = st.progress(0)
    status_text = st.empty()
    st.info("⏳ Elaborazione in corso...")

    att = pd.read_excel(file_att)
    tab_raw = pd.read_excel(file_tab, header=None, skiprows=3)
    tab_raw.columns = tab_raw.iloc[0]
    tab = tab_raw.drop(0).reset_index(drop=True)
    tab = tab.rename(columns={"macroarea": "Macroarea"})

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
    totale = len(tab)

    for i, (_, r) in enumerate(tab.iterrows(), 1):
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

        if i % 20 == 0 or i == totale:
            progress_bar.progress(i/totale)
            status_text.text(f"Elaborazione clienti: {int(i/totale*100)}%")

    status_text.text("Analisi attività senza cliente...")

    clienti_norm = set(tab["Cliente_n"].dropna().unique())
    att_no_match = att[~att["NomeSoggetto_n"].isin(clienti_norm)].copy()

    if not att_no_match.empty:
        att_no_match = (
            att_no_match.sort_values(["Anno", "Mese", "Priorita"])
            .groupby("NomeSoggetto", as_index=False)
            .last()
        )

        def da_ria_att(row):
            anno = int(row["Anno"])
            mese = int(row["Mese"])
            diff = (2025 - anno) * 12 + (11 - mese)
            return "Sì" if diff > 2 else "No"

        att_no_match["Da riassegnare"] = att_no_match.apply(da_ria_att, axis=1)
        att_no_match["Sede"] = att_no_match["Sede"]
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

    # 6️⃣ Esporta Excel + formattazione
    database = pd.DataFrame(righe_output).replace({np.nan: ""})

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

    for c in ["PREVENTIVATO€","DELIBERATO€","FATTURATO€","INCASSATO€"]:
        if c in database.columns:
            database[c] = database[c].apply(to_float_euro).apply(format_euro)

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

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    elapsed = time.time() - start_time
    minuti = int(elapsed // 60)
    secondi = int(elapsed % 60)

    progress_bar.empty()
    status_text.empty()

    st.success(f"✅ Report completato in {minuti} min {secondi} sec!")
    st.download_button(
        label="📥 Scarica report_attivita_clienti.xlsx",
        data=buf,
        file_name="report_attivita_clienti.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ======================
# 6️⃣ FOOTER
# ======================
st.markdown("""
---
<div style='text-align:center; font-size:13px; color:gray;'>
© 2025 <b>EdiliziAcrobatica S.p.A.</b> – Uso interno riservato<br>
<a href='https://www.ediliziacrobatica.com' target='_blank' style='color:#004C97; text-decoration:none;'>www.ediliziacrobatica.com</a>
</div>
""", unsafe_allow_html=True)
