import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ========================================
# CONFIGURAZIONE BASE
# ========================================
st.set_page_config(
    page_title="📊 Report Attività Clienti",
    page_icon="📘",
    layout="centered"
)

st.title("📊 Generatore Report Attività Clienti")

st.markdown("""
Carica i due file Excel generati dalla **Dashboard Commerciale** per ottenere automaticamente il report completo con attività, clienti e assegnazioni.

---
""")

# ========================================
# SEZIONE FILE ATTIVITÀ
# ========================================
st.markdown("""
### 📘 File Attività

📍 **Dove trovarlo:**  
Dashboard Commerciale → **Sottoprodotti → Tab Grafici Attività**  
Seleziona **l’ultimo elenco prima del grafico “Delibere”**  
➡️ Attendi il caricamento dei dati e premi **Crea Excel**
""")

file_att = st.file_uploader("📤 Seleziona il file delle attività (.xlsx)", type=["xlsx"], key="attivita")

st.divider()

# ========================================
# SEZIONE FILE CLIENTI
# ========================================
st.markdown("""
### 📗 File Clienti

📍 **Dove trovarlo:**  
Dashboard Commerciale → **Riepilogo Clienti**  
Imposta il periodo **dal 2017 ad oggi**,  
e scarica Excel da **“Tabella Clienti (no filtro data)”** in fondo alla pagina,  
dopo aver atteso il caricamento dei dati.
""")

file_cli = st.file_uploader("📤 Seleziona la tabella clienti (.xlsx)", type=["xlsx"], key="clienti")

st.divider()

# ========================================
# ELABORAZIONE FILE
# ========================================
if file_att and file_cli:
    st.success("✅ File caricati correttamente. Pronto a generare il report.")

    if st.button("🚀 Genera report Excel"):
        with st.spinner("⏳ Elaborazione in corso..."):

            # ----------------------------
            # LETTURA FILE CLIENTI
            # ----------------------------
            tab = pd.read_excel(file_cli, header=None, skiprows=3)
            tab.columns = tab.iloc[0]
            tab = tab.drop(0).reset_index(drop=True)
            tab = tab.rename(columns={c: str(c).strip() for c in tab.columns})

            # LETTURA FILE ATTIVITÀ
            att = pd.read_excel(file_att)
            att.columns = [c.strip() for c in att.columns]

            # ----------------------------
            # NORMALIZZAZIONE NOMI
            # ----------------------------
            def normalize(x):
                if pd.isna(x): return ""
                x = str(x).lower().replace(".", " ").replace("*", " ").replace(",", " ")
                return " ".join(x.split())

            att["NomeSoggetto_n"] = att["NomeSoggetto"].apply(normalize)
            tab["Cliente_n"] = tab["Cliente"].apply(normalize)

            # ----------------------------
            # PRIORITÀ ATTIVITÀ
            # ----------------------------
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

            # ----------------------------
            # MATCH E RIEPILOGO
            # ----------------------------
            risultati = []

            for _, r in tab.iterrows():
                cliente_norm = r["Cliente_n"]
                tipo = r.get("Tipo", "Amministratori")
                sede = r.get("Sede", "")
                resp = r.get("Responsabile", "")

                att_cli = att[att["NomeSoggetto_n"] == cliente_norm]
                if att_cli.empty:
                    invertito = " ".join(cliente_norm.split()[::-1])
                    att_cli = att[att["NomeSoggetto_n"] == invertito]

                if not att_cli.empty:
                    att_sel = att_cli.sort_values(["Anno", "Mese", "Priorita"]).iloc[-1]
                    anno, mese = int(att_sel["Anno"]), int(att_sel["Mese"])
                    diff = (2025 - anno) * 12 + (11 - mese)
                    da_ria = "Sì" if diff > 2 else "No"
                    risultati.append({
                        "Sede": sede,
                        "Responsabile gestionale": resp,
                        "Cliente": r["Cliente"],
                        "Anno": anno,
                        "Mese": mese,
                        "Ultima attività": att_sel["Classe Attività"],
                        "Da riassegnare": da_ria,
                        "PREVENTIVATO€": r.get("PREVENTIVATO€", ""),
                        "DELIBERATO€": r.get("DELIBERATO€", ""),
                        "FATTURATO€": r.get("FATTURATO€", ""),
                        "INCASSATO€": r.get("INCASSATO€", ""),
                        "Tipo": tipo
                    })
                else:
                    risultati.append({
                        "Sede": sede,
                        "Responsabile gestionale": resp,
                        "Cliente": r["Cliente"],
                        "Anno": "",
                        "Mese": "",
                        "Ultima attività": "",
                        "Da riassegnare": "Sì",
                        "PREVENTIVATO€": r.get("PREVENTIVATO€", ""),
                        "DELIBERATO€": r.get("DELIBERATO€", ""),
                        "FATTURATO€": r.get("FATTURATO€", ""),
                        "INCASSATO€": r.get("INCASSATO€", ""),
                        "Tipo": tipo
                    })

            df = pd.DataFrame(risultati).replace({np.nan: ""})

            # ----------------------------
            # FORMATTAZIONE NUMERI
            # ----------------------------
            def format_euro(x):
                if x == "" or pd.isna(x): return ""
                try:
                    val = float(str(x).replace("€", "").replace(".", "").replace(",", "."))
                    return f"€ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                except:
                    return str(x)

            for c in ["PREVENTIVATO€","DELIBERATO€","FATTURATO€","INCASSATO€"]:
                if c in df.columns:
                    df[c] = df[c].apply(format_euro)

            order = ["Sede","Responsabile gestionale","Cliente","Anno","Mese",
                     "Ultima attività","Da riassegnare","PREVENTIVATO€","DELIBERATO€","FATTURATO€","INCASSATO€"]

            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                df.to_excel(w, sheet_name="Database", index=False)
                for tipo, g in sorted(df.groupby("Tipo"), key=lambda x: str(x[0])):
                    name = str(tipo).capitalize() or "Senzatipo"
                    g[order].sort_values("Cliente").to_excel(w, sheet_name=name, index=False)

            wb = load_workbook(buf)
            thin = Side(border_style="thin", color="D9D9D9")
            red_fill = PatternFill(start_color="FFB6B6", end_color="FFB6B6", fill_type="solid")
            green_fill = PatternFill(start_color="B7E1B0", end_color="B7E1B0", fill_type="solid")
            blue_fill = PatternFill(start_color="004C97", end_color="004C97", fill_type="solid")

            for ws in wb.worksheets:
                ws.auto_filter.ref = ws.dimensions
                for c in ws[1]:
                    c.font = Font(bold=True, color="FFFFFF")
                    c.fill = blue_fill
                    c.alignment = Alignment(horizontal="center")
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    for cell in row:
                        if cell.value == "Sì": cell.fill = red_fill
                        elif cell.value == "No": cell.fill = green_fill
                        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

                # 🔹 AUTO-LARGHEZZA COLONNE
                for col_cells in ws.columns:
                    max_len = max(len(str(c.value)) if c.value else 0 for c in col_cells)
                    ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 45)

            # 🔹 APERTURA AUTOMATICA SUL FOGLIO "Amministratori"
            if "Amministratori" in wb.sheetnames:
                wb.active = wb.sheetnames.index("Amministratori")

            out = BytesIO()
            wb.save(out)
            out.seek(0)

            st.success("✅ File Excel generato con successo!")
            st.download_button(
                label="⬇️ Scarica il file Excel",
                data=out,
                file_name="report_attivita_clienti.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
