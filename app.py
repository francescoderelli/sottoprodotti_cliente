import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from tqdm import tqdm
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# === FUNZIONE PRINCIPALE ===
def genera_report(file_attivita, file_clienti):
    # Lettura file
    attivita = pd.read_excel(file_attivita)
    tabella = pd.read_excel(file_clienti, header=None, skiprows=3)
    tabella.columns = tabella.iloc[0]
    tabella = tabella.drop(0).reset_index(drop=True)
    tabella = tabella.rename(columns=lambda x: str(x).strip())
    attivita.columns = attivita.columns.str.strip()

    # Normalizza nomi
    def normalize_name(x):
        if pd.isna(x): return ""
        x = str(x).lower().replace(".", " ").replace("*", " ").replace(",", " ")
        return " ".join(x.split())
    tabella["Cliente_n"] = tabella["Cliente"].apply(normalize_name)
    attivita["NomeSoggetto_n"] = attivita["NomeSoggetto"].apply(normalize_name)

    # Normalizza tipo
    if "Tipo" in tabella.columns:
        def fix_tipo(x):
            x = str(x).strip().capitalize()
            if x.lower().startswith("amministrator"):
                return "Amministratori"
            return x
        tabella["Tipo"] = tabella["Tipo"].apply(fix_tipo)
    else:
        tabella["Tipo"] = "Amministratori"

    # Priorità attività
    priorita = {
        "04 RICHIESTE": 1, "06 PREVENTIVI": 2, "03 INCONTRI": 3,
        "07 DELIBERE": 4, "05 SOPRALLUOGHI": 5, "01 TELEFONATE": 6, "02 APPUNTAMENTI": 7
    }
    attivita["Priorita"] = attivita["Classe Attività"].map(priorita).fillna(999)

    # Matching clienti ↔ attività
    match_rows = []
    progress = st.progress(0)
    for i, (_, row) in enumerate(tabella.iterrows()):
        progress.progress(int((i+1)/len(tabella)*100))
        cliente = row["Cliente_n"]
        tipo = row["Tipo"]
        responsabile_gest = row.get("Responsabile", "")
        sede = row.get("Sede", "")
        match = attivita[attivita["NomeSoggetto_n"] == cliente]
        if match.empty:
            invertito = " ".join(cliente.split()[::-1])
            match = attivita[attivita["NomeSoggetto_n"] == invertito]
        if not match.empty:
            match = match.sort_values(by=["Anno", "Mese", "Priorita"]).iloc[-1]
            diff = (2025 - int(match["Anno"])) * 12 + (11 - int(match["Mese"]))
            da_ria = "Sì" if diff > 2 else "No"
            match_rows.append({
                "Sede": sede,
                "Responsabile gestionale": responsabile_gest,
                "Cliente": row["Cliente"],
                "Anno": match["Anno"],
                "Mese": match["Mese"],
                "Ultima attività": match["Classe Attività"],
                "Da riassegnare": da_ria,
                "PREVENTIVATO€": row.get("PREVENTIVATO€", ""),
                "DELIBERATO€": row.get("DELIBERATO€", ""),
                "FATTURATO€": row.get("FATTURATO€", ""),
                "INCASSATO€": row.get("INCASSATO€", ""),
                "Tipo": tipo
            })
        else:
            match_rows.append({
                "Sede": sede,
                "Responsabile gestionale": responsabile_gest,
                "Cliente": row["Cliente"],
                "Anno": "",
                "Mese": "",
                "Ultima attività": "",
                "Da riassegnare": "Sì",
                "PREVENTIVATO€": row.get("PREVENTIVATO€", ""),
                "DELIBERATO€": row.get("DELIBERATO€", ""),
                "FATTURATO€": row.get("FATTURATO€", ""),
                "INCASSATO€": row.get("INCASSATO€", ""),
                "Tipo": tipo
            })

    clienti_tabella = set(tabella["Cliente_n"].dropna().unique())
    no_match = attivita[~attivita["NomeSoggetto_n"].isin(clienti_tabella)].copy()
    if not no_match.empty:
        no_match_grouped = (
            no_match.sort_values(["Anno", "Mese", "Priorita"])
            .groupby("NomeSoggetto", as_index=False)
            .last()
        )
        def da_riassegnare(row):
            anno, mese = int(row["Anno"]), int(row["Mese"])
            diff = (2025 - anno) * 12 + (11 - mese)
            return "Sì"  # Tutti da riassegnare
        no_match_grouped["Da riassegnare"] = no_match_grouped.apply(da_riassegnare, axis=1)
        no_match_grouped["Sede"] = no_match_grouped["Sede"]
        no_match_grouped["Responsabile gestionale"] = no_match_grouped["Responsabile"]
        no_match_grouped["Cliente"] = no_match_grouped["NomeSoggetto"]
        no_match_grouped["Ultima attività"] = no_match_grouped["Classe Attività"]
        no_match_grouped["Tipo"] = "Amministratori"
        for col in ["PREVENTIVATO€", "DELIBERATO€", "FATTURATO€", "INCASSATO€"]:
            no_match_grouped[col] = ""
    else:
        no_match_grouped = pd.DataFrame()

    database = pd.DataFrame(match_rows)
    if not no_match_grouped.empty:
        database = pd.concat([database, no_match_grouped[database.columns]], ignore_index=True)

    # Format euro
    def format_euro(x):
        try:
            val = float(str(x).replace(",", ".").replace("€", "").strip())
            return f"€ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return ""
    for col in ["PREVENTIVATO€", "DELIBERATO€", "FATTURATO€", "INCASSATO€"]:
        database[col] = database[col].apply(format_euro)

    # Output Excel in memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        database.to_excel(writer, sheet_name="Database", index=False)
        col_order = [
            "Sede", "Responsabile gestionale", "Cliente", "Anno", "Mese",
            "Ultima attività", "Da riassegnare", "PREVENTIVATO€", "DELIBERATO€",
            "FATTURATO€", "INCASSATO€"
        ]
        for tipo, grp in sorted(database.groupby("Tipo"), key=lambda x: x[0]):
            grp = grp[col_order].sort_values("Cliente")
            grp.to_excel(writer, sheet_name=tipo.capitalize(), index=False)
    return output.getvalue()

# === INTERFACCIA STREAMLIT ===
st.set_page_config(page_title="Generatore Report Attività", page_icon="📊")
st.image("https://www.ediliziacrobatica.com/wp-content/uploads/2022/05/logo.svg", width=250)
st.title("🏗️ Generatore Report Attività Clienti")

file_att = st.file_uploader("📘 Carica il file attività (attivita_2025.xlsx)", type=["xlsx"])
file_tab = st.file_uploader("📗 Carica il file tabella clienti (Tabella_Clienti.xlsx)", type=["xlsx"])

if file_att and file_tab:
    if st.button("🚀 Crea file di output"):
        with st.spinner("Elaborazione in corso, attendere..."):
            result = genera_report(file_att, file_tab)
        st.success("✅ File generato con successo!")
        st.download_button(
            label="⬇️ Scarica file Excel",
            data=result,
            file_name="output_attivita_finale.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
