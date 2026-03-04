"""
Script de génération des fichiers Excel de test et de comparaison.
Génère security_ids.xlsx et security_refprod.xlsx dans data/,
puis produit le rapport de comparaison rapport_comparaison.md.
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import datetime

# ─── 1. CREATION DES FICHIERS EXCEL ───────────────────────────────────────────

def create_security_ids():
    data = {
        "ID_INTERNE": [
            "SEC001", "SEC002", "SEC003", "SEC004", "SEC005",
            "SEC006", "SEC007", "SEC008", "SEC009", "SEC010",
            "SEC011", "SEC012", "SEC013",
        ],
        "ISIN": [
            "FR0000131104", "US0378331005", "DE0005140008",
            "FR0000120271", "US5949181045", "GB0031348658",
            "FR0000120073", "US4592001014", "DE000BASF111",
            "FR0000120628", "US2546871060", "GB00B03MLX29",
            "FR0010208488",
        ],
        "CUSIP": [
            None, "037833100", None, None, "594918104",
            None, None, "459200101", None, None, "254687106",
            None, None,
        ],
        "SEDOL": [
            "5752815", "2046251", "5097905", "5809036", "2588173",
            "3134865", "5705928", "2452743", "5052188", "7012897",
            "2174541", "B03MLX2", "B0159S7",
        ],
        "RIC": [
            "AIRP.PA", "AAPL.O", "DBKGn.DE", "AIR.PA", "MSFT.O",
            "SHEL.L", "AIRP.PA", "IBM.N", "BASFn.DE", "OR.PA",
            "DIS.N", "SHEL.L", "AXAF.PA",
        ],
        "LIBELLE": [
            "AIR LIQUIDE SA", "APPLE INC", "DEUTSCHE BANK AG",
            "AIRBUS SE", "MICROSOFT CORP", "SHELL PLC",
            "AIR LIQUIDE SA (PREFX)", "IBM CORP", "BASF SE",
            "LOREAL SA", "WALT DISNEY CO", "SHELL PLC (ADR)",
            "AXA SA",
        ],
        "CLASSE_ACTIF": [
            "ACTION", "ACTION", "ACTION",
            "ACTION", "ACTION", "ACTION",
            "ACTION", "ACTION", "ACTION",
            "ACTION", "ACTION", "ACTION",
            "ACTION",
        ],
        "MARCHE": [
            "EURONEXT PARIS", "NASDAQ", "XETRA",
            "EURONEXT PARIS", "NASDAQ", "LSE",
            "EURONEXT PARIS", "NYSE", "XETRA",
            "EURONEXT PARIS", "NYSE", "NYSE",
            "EURONEXT PARIS",
        ],
        "DATE_CREATION": [
            "2020-01-15", "2019-06-10", "2021-03-22",
            "2020-07-01", "2018-11-05", "2022-02-14",
            "2020-01-15", "2017-09-30", "2021-08-19",
            "2019-04-25", "2020-12-01", "2022-02-14",
            "2023-01-10",
        ],
        "STATUT": [
            "ACTIF", "ACTIF", "ACTIF",
            "ACTIF", "ACTIF", "ACTIF",
            "INACTIF", "ACTIF", "ACTIF",
            "ACTIF", "ACTIF", "INACTIF",
            "ACTIF",
        ],
    }
    return pd.DataFrame(data)


def create_security_refprod():
    data = {
        "ID_INTERNE": [
            "SEC001", "SEC002", "SEC003", "SEC004", "SEC005",
            "SEC006", "SEC007", "SEC008", "SEC009", "SEC010",
            "SEC014", "SEC015", "SEC016",
        ],
        "ISIN": [
            "FR0000131104", "US0378331005", "DE0005140008",
            "FR0000120271", "US5949181045", "GB0031348658",
            "FR0000120073", "US4592001014", "DE000BASF111",
            "FR0000120628", "US38141G1040", "FR0013506730",
            "DE0005552004",
        ],
        "TYPE_PRODUIT": [
            "EQUITE", "EQUITE", "EQUITE",
            "EQUITE", "EQUITE", "EQUITE",
            "EQUITE", "EQUITE", "EQUITE",
            "EQUITE", "EQUITE", "OBLIGATION",
            "EQUITE",
        ],
        "EMETTEUR": [
            "AIR LIQUIDE", "APPLE INC", "DEUTSCHE BANK",
            "AIRBUS", "MICROSOFT", "SHELL",
            "AIR LIQUIDE", "IBM", "BASF",
            "LOREAL", "GOOGLE LLC", "FRANCE TRESOR",
            "BMW AG",
        ],
        "DEVISE": [
            "EUR", "USD", "EUR",
            "EUR", "USD", "GBP",
            "EUR", "USD", "EUR",
            "EUR", "USD", "EUR",
            "EUR",
        ],
        "PAYS": [
            "France", "États-Unis", "Allemagne",
            "France", "États-Unis", "Royaume-Uni",
            "France", "États-Unis", "Allemagne",
            "France", "États-Unis", "France",
            "Allemagne",
        ],
        "NOTATION": [
            "A+", "AA+", "BB",
            "BBB+", "AAA", "A",
            "A+", "A-", "A",
            "AA", None, "AA+",
            "A",
        ],
        "DATE_MISE_A_JOUR": [
            "2024-11-01", "2024-10-15", "2024-09-30",
            "2024-11-01", "2024-10-20", "2024-08-05",
            "2024-11-01", "2024-07-18", "2024-10-01",
            "2024-11-10", "2024-06-22", "2024-05-15",
            "2024-09-12",
        ],
        "MARCHE": [
            "EURONEXT PARIS", "NASDAQ", "XETRA",
            "EURONEXT PARIS", "NASDAQ", "LSE",
            "EURONEXT PARIS", "NYSE", "XETRA",
            "EURONEXT PARIS", "NASDAQ", "EURONEXT PARIS",
            "XETRA",
        ],
        "STATUT": [
            "ACTIF", "ACTIF", "ACTIF",
            "ACTIF", "ACTIF", "ACTIF",
            "ACTIF", "ACTIF", "ACTIF",
            "ACTIF", "ACTIF", "ACTIF",
            "ACTIF",
        ],
    }
    return pd.DataFrame(data)


def style_excel(ws, df, header_color="1F4E79"):
    """Apply styling to an Excel worksheet."""
    header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    alt_fill = PatternFill(start_color="EBF3FB", end_color="EBF3FB", fill_type="solid")
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border
            cell.alignment = Alignment(vertical="center")
            if row_idx % 2 == 0:
                cell.fill = alt_fill

    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col) + 4
        ws.column_dimensions[col[0].column_letter].width = min(max_len, 30)
    ws.row_dimensions[1].height = 22


def save_excel(df, path, sheet_name, header_color="1F4E79"):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)
    style_excel(ws, df, header_color)
    wb.save(path)
    print(f"  ✔ Fichier créé : {path}")


# ─── 2. COMPARAISON ──────────────────────────────────────────────────────────

def compare_dataframes(df_ids, df_ref):
    """Perform a thorough comparison between the two DataFrames."""
    results = {}

    # 2.1 Structure
    results["cols_ids"] = set(df_ids.columns)
    results["cols_ref"] = set(df_ref.columns)
    results["cols_common"] = results["cols_ids"] & results["cols_ref"]
    results["cols_only_ids"] = results["cols_ids"] - results["cols_ref"]
    results["cols_only_ref"] = results["cols_ref"] - results["cols_ids"]

    results["rows_ids"] = len(df_ids)
    results["rows_ref"] = len(df_ref)

    # 2.2 Row sets (key = ID_INTERNE)
    key = "ID_INTERNE"
    ids_set = set(df_ids[key])
    ref_set = set(df_ref[key])
    results["ids_common_keys"] = ids_set & ref_set
    results["ids_only_ids"] = ids_set - ref_set
    results["ids_only_ref"] = ref_set - ids_set

    # 2.3 Common rows – column-level diff
    df_ids_common = df_ids[df_ids[key].isin(results["ids_common_keys"])].set_index(key)
    df_ref_common = df_ref[df_ref[key].isin(results["ids_common_keys"])].set_index(key)
    common_cols = sorted(results["cols_common"] - {key})

    diff_details = {}
    for col in common_cols:
        if col not in df_ids_common.columns or col not in df_ref_common.columns:
            continue
        col_diffs = []
        for idx in sorted(results["ids_common_keys"]):
            val_ids = df_ids_common.at[idx, col] if idx in df_ids_common.index else None
            val_ref = df_ref_common.at[idx, col] if idx in df_ref_common.index else None
            if pd.isna(val_ids) and pd.isna(val_ref):
                continue
            if str(val_ids) != str(val_ref):
                col_diffs.append((idx, val_ids, val_ref))
        if col_diffs:
            diff_details[col] = col_diffs

    results["diff_details"] = diff_details

    # 2.4 ISIN coherence check
    ids_isin = set(df_ids["ISIN"])
    ref_isin = set(df_ref["ISIN"])
    results["isin_common"] = ids_isin & ref_isin
    results["isin_only_ids"] = ids_isin - ref_isin
    results["isin_only_ref"] = ref_isin - ids_isin

    # 2.5 Per-column stats for common columns
    col_stats = {}
    for col in common_cols:
        if col not in df_ids.columns or col not in df_ref.columns:
            continue
        col_stats[col] = {
            "null_ids": int(df_ids[col].isna().sum()),
            "null_ref": int(df_ref[col].isna().sum()),
            "unique_ids": int(df_ids[col].nunique()),
            "unique_ref": int(df_ref[col].nunique()),
        }
    results["col_stats"] = col_stats

    return results


# ─── 3. GÉNÉRATION DU RAPPORT ─────────────────────────────────────────────────

def generate_report(r, path_ids, path_ref):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    lines = []

    def h1(t):  lines.append(f"\n# {t}\n")
    def h2(t):  lines.append(f"\n## {t}\n")
    def h3(t):  lines.append(f"\n### {t}\n")
    def para(t): lines.append(f"{t}\n")
    def sep():  lines.append("---\n")

    h1("Rapport de Comparaison des Fichiers Excel")
    para(f"**Date de génération :** {now}  ")
    para(f"**Fichier 1 (IDS) :** `{os.path.basename(path_ids)}`  ")
    para(f"**Fichier 2 (REFPROD) :** `{os.path.basename(path_ref)}`  ")
    sep()

    # ── Section 1 : Résumé exécutif
    h2("1. Résumé Exécutif")
    nb_common_keys = len(r["ids_common_keys"])
    nb_diff_cols = len(r["diff_details"])
    para(f"- **Lignes dans security_ids :** {r['rows_ids']}")
    para(f"- **Lignes dans security_refprod :** {r['rows_ref']}")
    para(f"- **Lignes communes (même ID_INTERNE) :** {nb_common_keys}")
    para(f"- **Lignes uniques à security_ids :** {len(r['ids_only_ids'])}")
    para(f"- **Lignes uniques à security_refprod :** {len(r['ids_only_ref'])}")
    para(f"- **Colonnes dans security_ids :** {len(r['cols_ids'])}")
    para(f"- **Colonnes dans security_refprod :** {len(r['cols_ref'])}")
    para(f"- **Colonnes communes :** {len(r['cols_common'])}")
    para(f"- **Colonnes exclusives à security_ids :** {len(r['cols_only_ids'])}")
    para(f"- **Colonnes exclusives à security_refprod :** {len(r['cols_only_ref'])}")
    para(f"- **Colonnes communes avec divergences de valeur :** {nb_diff_cols}")

    # ── Section 2 : Structure des fichiers
    h2("2. Structure des Fichiers")

    h3("2.1 Colonnes de security_ids")
    lines.append("| # | Colonne | Présente dans REFPROD |\n|---|---------|----------------------|\n")
    for i, c in enumerate(sorted(r["cols_ids"]), 1):
        tag = "✅ Oui" if c in r["cols_ref"] else "❌ Non"
        lines.append(f"| {i} | `{c}` | {tag} |\n")

    h3("2.2 Colonnes de security_refprod")
    lines.append("| # | Colonne | Présente dans IDS |\n|---|---------|-------------------|\n")
    for i, c in enumerate(sorted(r["cols_ref"]), 1):
        tag = "✅ Oui" if c in r["cols_ids"] else "❌ Non"
        lines.append(f"| {i} | `{c}` | {tag} |\n")

    h3("2.3 Colonnes exclusives")
    if r["cols_only_ids"]:
        para(f"**Uniquement dans security_ids :** {', '.join(f'`{c}`' for c in sorted(r['cols_only_ids']))}")
    if r["cols_only_ref"]:
        para(f"**Uniquement dans security_refprod :** {', '.join(f'`{c}`' for c in sorted(r['cols_only_ref']))}")
    if not r["cols_only_ids"] and not r["cols_only_ref"]:
        para("*Aucune colonne exclusive : les deux fichiers partagent exactement les mêmes colonnes.*")

    # ── Section 3 : Analyse des lignes
    h2("3. Analyse des Lignes (clé : ID_INTERNE)")

    h3("3.1 Lignes communes")
    para(f"{nb_common_keys} identifiants présents dans les deux fichiers :")
    lines.append("\n" + ", ".join(f"`{k}`" for k in sorted(r["ids_common_keys"])) + "\n")

    h3("3.2 Lignes uniques à security_ids")
    if r["ids_only_ids"]:
        para(f"{len(r['ids_only_ids'])} ligne(s) absente(s) de security_refprod :")
        lines.append("\n" + ", ".join(f"`{k}`" for k in sorted(r["ids_only_ids"])) + "\n")
    else:
        para("*Aucune ligne unique dans security_ids.*")

    h3("3.3 Lignes uniques à security_refprod")
    if r["ids_only_ref"]:
        para(f"{len(r['ids_only_ref'])} ligne(s) absente(s) de security_ids :")
        lines.append("\n" + ", ".join(f"`{k}`" for k in sorted(r["ids_only_ref"])) + "\n")
    else:
        para("*Aucune ligne unique dans security_refprod.*")

    # ── Section 4 : Cohérence ISIN
    h2("4. Cohérence des Codes ISIN")
    para(f"- ISIN communs aux deux fichiers : **{len(r['isin_common'])}**")
    para(f"- ISIN uniquement dans security_ids : **{len(r['isin_only_ids'])}**")
    para(f"- ISIN uniquement dans security_refprod : **{len(r['isin_only_ref'])}**")

    if r["isin_only_ids"]:
        h3("ISIN présents uniquement dans security_ids")
        for isin in sorted(r["isin_only_ids"]):
            lines.append(f"- `{isin}`\n")

    if r["isin_only_ref"]:
        h3("ISIN présents uniquement dans security_refprod")
        for isin in sorted(r["isin_only_ref"]):
            lines.append(f"- `{isin}`\n")

    # ── Section 5 : Divergences de valeur
    h2("5. Divergences de Valeur sur les Colonnes Communes")

    if not r["diff_details"]:
        para("✅ Aucune divergence de valeur détectée pour les colonnes communes sur les lignes communes.")
    else:
        para(f"{len(r['diff_details'])} colonne(s) présentent des différences de valeur :")
        for col, diffs in sorted(r["diff_details"].items()):
            h3(f"Colonne : `{col}`")
            lines.append(f"| ID_INTERNE | Valeur dans IDS | Valeur dans REFPROD |\n")
            lines.append(f"|------------|-----------------|---------------------|\n")
            for idx, v1, v2 in diffs:
                v1_str = str(v1) if not (isinstance(v1, float) and np.isnan(v1)) else "*(vide)*"
                v2_str = str(v2) if not (isinstance(v2, float) and np.isnan(v2)) else "*(vide)*"
                lines.append(f"| `{idx}` | {v1_str} | {v2_str} |\n")

    # ── Section 6 : Statistiques par colonne
    h2("6. Statistiques par Colonne Commune")
    lines.append("| Colonne | Valeurs nulles IDS | Valeurs nulles REFPROD | Valeurs uniques IDS | Valeurs uniques REFPROD |\n")
    lines.append("|---------|-------------------|------------------------|---------------------|-------------------------|\n")
    for col, s in sorted(r["col_stats"].items()):
        lines.append(
            f"| `{col}` | {s['null_ids']} | {s['null_ref']} | {s['unique_ids']} | {s['unique_ref']} |\n"
        )

    # ── Section 7 : Conclusions
    h2("7. Conclusions et Points de Divergence")
    sep()

    h3("Points de Similitude")
    lines.append("- Les deux fichiers utilisent `ID_INTERNE` et `ISIN` comme identifiants clés.\n")
    lines.append(f"- {nb_common_keys} valeurs d'`ID_INTERNE` sont partagées entre les deux fichiers.\n")
    common_cols_shared = sorted(r["cols_common"])
    lines.append(f"- {len(common_cols_shared)} colonne(s) sont communes : {', '.join(f'`{c}`' for c in common_cols_shared)}.\n")
    lines.append("- La colonne `MARCHE` et `STATUT` apparaissent dans les deux fichiers avec des valeurs globalement cohérentes.\n")

    h3("Points de Divergence")
    if r["cols_only_ids"]:
        lines.append(f"- **Colonnes IDS uniquement :** {', '.join(f'`{c}`' for c in sorted(r['cols_only_ids']))} — ces colonnes d'identification ne sont pas portées par REFPROD.\n")
    if r["cols_only_ref"]:
        lines.append(f"- **Colonnes REFPROD uniquement :** {', '.join(f'`{c}`' for c in sorted(r['cols_only_ref']))} — données de référence produit absentes de IDS.\n")
    if r["ids_only_ids"]:
        lines.append(f"- **{len(r['ids_only_ids'])} titre(s) présent(s) dans IDS mais absents de REFPROD** : {', '.join(sorted(r['ids_only_ids']))}. Ces titres n'ont pas encore de fiche produit de référence.\n")
    if r["ids_only_ref"]:
        lines.append(f"- **{len(r['ids_only_ref'])} titre(s) présent(s) dans REFPROD mais absents de IDS** : {', '.join(sorted(r['ids_only_ref']))}. Ces titres n'ont pas d'identifiants de marché associés.\n")
    if r["diff_details"]:
        lines.append(f"- **Divergences de valeur** sur {nb_diff_cols} colonne(s) pour les lignes communes : {', '.join(f'`{c}`' for c in sorted(r['diff_details']))}.\n")
    else:
        lines.append("- Aucune divergence de valeur détectée sur les colonnes et lignes communes.\n")

    h3("Recommandations")
    lines.append("1. **Aligner les périmètres** : les titres SEC011, SEC012, SEC013 (IDS) et SEC014, SEC015, SEC016 (REFPROD) doivent être réconciliés.\n")
    lines.append("2. **Compléter les fiches REFPROD** pour les titres présents uniquement dans IDS.\n")
    lines.append("3. **Enrichir IDS** avec les données d'identification manquantes pour les titres REFPROD.\n")
    lines.append("4. **Vérifier les divergences** de valeur colonne par colonne, notamment pour `STATUT`.\n")
    lines.append("5. **Mettre en place un processus de synchronisation** régulier entre les deux référentiels.\n")

    report_text = "".join(lines)
    report_path = os.path.join(os.path.dirname(path_ids), "..", "rapport_comparaison.md")
    report_path = os.path.normpath(report_path)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report_text)
    print(f"  ✔ Rapport généré : {report_path}")
    return report_path


# ─── MAIN ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    data_dir = os.path.join(os.path.dirname(__file__), "data")
    os.makedirs(data_dir, exist_ok=True)

    path_ids = os.path.join(data_dir, "security_ids.xlsx")
    path_ref = os.path.join(data_dir, "security_refprod.xlsx")

    print("\n=== Génération des fichiers Excel ===")
    df_ids = create_security_ids()
    df_ref = create_security_refprod()
    save_excel(df_ids, path_ids, "security_ids", header_color="1F4E79")
    save_excel(df_ref, path_ref, "security_refprod", header_color="375623")

    print("\n=== Comparaison en cours ===")
    r = compare_dataframes(df_ids, df_ref)

    print("\n=== Génération du rapport ===")
    report_path = generate_report(r, path_ids, path_ref)

    print("\n✅ Terminé. Fichiers produits :")
    print(f"   {path_ids}")
    print(f"   {path_ref}")
    print(f"   {report_path}")
