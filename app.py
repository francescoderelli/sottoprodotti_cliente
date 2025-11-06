import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ========================================
# CONFIGURAZIONE STREAMLIT
# ========================================
st.set_page_config(
    page_title="📊 Report Attività Clienti",
    page_icon="📘",
    layout="centered"
)

st.title("📊 Generatore report attività clienti")

st.markdown("""
### 🔹 Carica i tuoi file Excel

**📁 File Attività**  
Scaricalo dalla **Dashboard Commerciale → Sottoprodotti → Tab Grafici Attività**,  
seleziona l’ultimo elenco prima del grafico *“Delibere”*.  
➡️ Attendi il caricamento dei dati e premi **Crea Excel**.

---

**📁 File Clienti**  
Scaricalo dalla **Dashboard Commerciale → Riepilogo Clienti**,  
impostando il periodo **dal 2017 ad oggi**,  
e scarica Excel da *“Tabella Clienti (no filtro data)”* in fondo alla pagina,  
dopo aver atteso il caricamento dei dati.
""")

# ========================================
# CARICAMENTO FILES
# ========================================
file1 = st.file_uploader("📂 Carica il primo file (.xlsx)", type=["xlsx"])
file2 = st.file_uploader("📂 Carica il secondo file (.xlsx)", type=["xlsx"])

if file1 and file2:
    st.success("✅ File caricati con successo! Ora rilevo automaticamente quale è Attività e quale è Tabella Clienti.")
    if st.button("🚀 Genera report Excel"):
        with st.spinner("⏳ Elaborazione in corso..."):

            # ----------------------------------------
            # 1️⃣ Tenta di capire quale file è quale
            # ----------------------------------------
            def leggi_file_auto(file):
                try:
                    df = pd.read_excel(file)
                    cols = [c.lower() for c in df.columns]
                    if any("classe attività" in c for c in cols):
                        return "attivita", df
                except Exception:
                    pass

                try:
                    df = pd.read_excel(file, header=None, skiprows=3)
                    df.columns = df.iloc[0]
                    df = df.drop(0).reset_index(drop=True)
                    cols = [str(c).lower() for c in df.columns]
                    if any("cliente" in c for c in cols) or "tabella clienti" in cols[0]:
                        return "clienti", df
                except Exception:
                    pass

                return "sconosciuto", None

            tipo1, df1 = leggi_file_auto(file1)
            tipo2, df2 = leggi_file_auto(file2)

            if tipo1 == "attivita" and tipo2 == "clienti":
                att, tab = df1, df2
            elif tipo1 == "clienti" and tipo2 == "attivita":
                att, tab = df2, df1
            else:
                st.error("⚠️ Non riesco a riconoscere i file. Assicurati di caricare un file Attività e uno Clienti.")
                st.stop()

            # ----------------------------------------
            # 2️⃣ Controllo colonne minime
            # ----------------------------------------
            colonne_attese_att = ["Anno","Mese","Classe Attività","Responsabile","NomeSoggetto"]
            colonne_attese_tab = ["Cliente","Sede","Responsabile"]

            if not all(any(col.lower() == c.lower() for c in att.columns) for col in colonne_attese_att):
                st.error("⚠️ Il file Attività caricato non contiene tutte le colonne richieste.")
                st.stop()

            if not any("cliente" in str(c).lower() for c in tab.columns):
                st.error("⚠️ Il file Tabella Clienti non sembra corretto.")
                st.stop()

            # ----------------------------------------
            # 3️⃣ Normalizzazione nomi
            # ----------------------------------------
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

            # ----------------------------------------
            # 4️⃣ Match tra tabella e attività
            # ----------------------------------------
            righe_output = []

            for _, r in tab.iterrows():
                cliente_norm = r["Cliente_n"]
                tipo_cli = r.get("Tipo", "Amministratori")
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

            database = pd.DataFrame(righe_output).replace({np.nan: ""})

            # ----------------------------------------
            # 5️⃣ Formatting numerico
            # ----------------------------------------
            def format_euro(x):
                if x == "" or pd.isna(x): return ""
                try:
                    val = float(str(x).replace("€","").replace(".","").replace(",","."))
                    return f"€ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                except:
                    return str(x)

            for c in ["PREVENTIVATO€","DELIBERATO€","FATTURATO€","INCASSATO€"]:
                if c in database.columns:
                    database[c] = database[c].apply(format_euro)

            # ----------------------------------------
            # 6️⃣ Generazione Excel con formattazione
            # ----------------------------------------
            col_order = [
                "Sede","Responsabile gestionale","Cliente","Anno","Mese",
                "Ultima attività","Da riassegnare",
                "PREVENTIVATO€","DELIBERATO€","FATTURATO€","INCASSATO€"
            ]

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                database.to_excel(writer, sheet_name="Database", index=False)
                for tipo, grp in sorted(database.groupby("Tipo"), key=lambda x: str(x[0])):
                    nome = str(tipo).strip().capitalize() or "Senzatipo"
                    grp[col_order].sort_values("Cliente").to_excel(writer, sheet_name=nome, index=False)

            wb = load_workbook(buffer)
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

            out_buffer = BytesIO()
            wb.save(out_buffer)
            out_buffer.seek(0)

            st.success("✅ File Excel generato con successo!")
            st.download_button(
                label="⬇️ Scarica il file Excel",
                data=out_buffer,
                file_name="report_attivita_clienti.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
