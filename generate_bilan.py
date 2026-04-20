import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# --- STYLES ---
title_font = Font(name='Calibri', bold=True, size=14)
header_font = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
section_font = Font(name='Calibri', bold=True, size=11)
subsection_font = Font(name='Calibri', bold=True, size=10)
normal_font = Font(name='Calibri', size=10)
total_font = Font(name='Calibri', bold=True, size=10)

header_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
section_fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
total_fill = PatternFill(start_color='B4C6E7', end_color='B4C6E7', fill_type='solid')
subtotal_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')

thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

euro_format = '#,##0.00 €'


def style_header_row(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin_border


def style_section_row(ws, row, cols, fill=section_fill, font=section_font):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = font
        cell.fill = fill
        cell.border = thin_border


def write_data_row(ws, row, label, v2023, v2024, v2025, indent=0, is_total=False):
    cell_a = ws.cell(row=row, column=1, value=(' ' * indent * 3) + label)
    cell_a.font = total_font if is_total else normal_font
    cell_a.border = thin_border
    cell_a.alignment = Alignment(wrap_text=True)

    for col_idx, val in enumerate([v2023, v2024, v2025], start=2):
        cell = ws.cell(row=row, column=col_idx, value=val)
        cell.number_format = euro_format
        cell.font = total_font if is_total else normal_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='right')

    if is_total:
        for c in range(1, 5):
            ws.cell(row=row, column=c).fill = total_fill


# ============================================================
# FEUILLE 1 : BILAN COMPARATIF
# ============================================================
ws_bilan = wb.active
ws_bilan.title = "Bilan Comparatif"
ws_bilan.column_dimensions['A'].width = 45
ws_bilan.column_dimensions['B'].width = 18
ws_bilan.column_dimensions['C'].width = 18
ws_bilan.column_dimensions['D'].width = 18

r = 1
ws_bilan.merge_cells('A1:D1')
cell = ws_bilan.cell(row=r, column=1, value="BILAN COMPARATIF - SDC 17 Rue Léon Cladel, Sèvres")
cell.font = title_font
cell.alignment = Alignment(horizontal='center')
r += 1
ws_bilan.merge_cells('A2:D2')
cell = ws_bilan.cell(row=r, column=1, value="Copropriété PERDUE - Syndic MANDA")
cell.font = Font(name='Calibri', italic=True, size=11)
cell.alignment = Alignment(horizontal='center')
r += 2

# Headers
ws_bilan.cell(row=r, column=1, value="ACTIF")
ws_bilan.cell(row=r, column=2, value="2023\n(31/12/2023)")
ws_bilan.cell(row=r, column=3, value="2024\n(31/12/2024)")
ws_bilan.cell(row=r, column=4, value="2025\n(20/07/2025)*")
style_header_row(ws_bilan, r, 4)
r += 1

# --- ACTIF ---
# Solde en attente travaux (1200)
style_section_row(ws_bilan, r, 4)
ws_bilan.cell(row=r, column=1, value="ACTIF IMMOBILISÉ")
r += 1
write_data_row(ws_bilan, r, "Solde en attente sur travaux (1200)", 120.00, 0, 0, indent=1)
r += 1
write_data_row(ws_bilan, r, "Total Actif immobilisé", 120.00, 0, 0, is_total=True)
r += 2

style_section_row(ws_bilan, r, 4)
ws_bilan.cell(row=r, column=1, value="ACTIF CIRCULANT")
r += 1
write_data_row(ws_bilan, r, "Copropriétaires - soldes débiteurs (450)", 0, 0, 3776.31, indent=1)
r += 1
write_data_row(ws_bilan, r, "Fournisseurs - soldes débiteurs (4010)", 0, 0, 278.48, indent=1)
r += 1
write_data_row(ws_bilan, r, "Régularisation charges - Débiteur (4711)", 18665.97, 26363.21, 26363.21, indent=1)
r += 1
write_data_row(ws_bilan, r, "Régularisation travaux - Débiteur (4712)", 0, 911.97, 911.97, indent=1)
r += 1
write_data_row(ws_bilan, r, "Rompus débiteurs (4730)", 0.06, 0.09, 0.09, indent=1)
r += 1
write_data_row(ws_bilan, r, "Total Actif circulant", 18666.03, 27275.27, 31330.06, is_total=True)
r += 2

style_section_row(ws_bilan, r, 4)
ws_bilan.cell(row=r, column=1, value="TRÉSORERIE ACTIVE")
r += 1
write_data_row(ws_bilan, r, "Compte courant Montepaschi (5120 0001)", 2968.03, 9525.90, 13074.46, indent=1)
r += 1
write_data_row(ws_bilan, r, "Livret A Monte Paschi (5120 0004)", 910.22, 1742.55, 1967.53, indent=1)
r += 1
write_data_row(ws_bilan, r, "Total Trésorerie", 3878.25, 11268.45, 15041.99, is_total=True)
r += 2

# TOTAL ACTIF
write_data_row(ws_bilan, r, "TOTAL ACTIF", 22664.28, 38543.72, 46372.05, is_total=True)
for c in range(1, 5):
    ws_bilan.cell(row=r, column=c).fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    ws_bilan.cell(row=r, column=c).font = Font(name='Calibri', bold=True, size=11)
r += 3

# --- PASSIF ---
ws_bilan.cell(row=r, column=1, value="PASSIF")
ws_bilan.cell(row=r, column=2, value="2023\n(31/12/2023)")
ws_bilan.cell(row=r, column=3, value="2024\n(31/12/2024)")
ws_bilan.cell(row=r, column=4, value="2025\n(20/07/2025)*")
style_header_row(ws_bilan, r, 4)
r += 1

style_section_row(ws_bilan, r, 4)
ws_bilan.cell(row=r, column=1, value="CAPITAUX PROPRES & AVANCES")
r += 1
write_data_row(ws_bilan, r, "Avances de trésorerie (1031)", 2000.00, 2000.00, 2000.00, indent=1)
r += 1
write_data_row(ws_bilan, r, "Fonds de travaux (1050)", 910.22, 1742.55, 2256.18, indent=1)
r += 1
write_data_row(ws_bilan, r, "Total Capitaux propres & Avances", 2910.22, 3742.55, 4256.18, is_total=True)
r += 2

style_section_row(ws_bilan, r, 4)
ws_bilan.cell(row=r, column=1, value="DETTES")
r += 1
write_data_row(ws_bilan, r, "Fournisseurs créditeurs (4010)", 1330.06, 659.45, 0, indent=1)
r += 1
write_data_row(ws_bilan, r, "Factures non parvenues (4080)", 0, 141.72, 0, indent=1)
r += 1
write_data_row(ws_bilan, r, "Tiers créditeurs - anciens copro (4630)", 424.00, 0, 0, indent=1)
r += 1
write_data_row(ws_bilan, r, "Régularisation charges - Créditeur (4721)", 18000.00, 34000.00, 34000.00, indent=1)
r += 1
write_data_row(ws_bilan, r, "Total Dettes", 19754.06, 34801.17, 34000.00, is_total=True)
r += 2

style_section_row(ws_bilan, r, 4)
ws_bilan.cell(row=r, column=1, value="RÉSULTAT EN COURS (exercice non clôturé)")
r += 1
write_data_row(ws_bilan, r, "Provisions appelées - Charges engagées", 0, 0, 8115.87, indent=1)
r += 1
write_data_row(ws_bilan, r, "Total Résultat en cours", 0, 0, 8115.87, is_total=True)
r += 2

# TOTAL PASSIF
write_data_row(ws_bilan, r, "TOTAL PASSIF", 22664.28, 38543.72, 46372.05, is_total=True)
for c in range(1, 5):
    ws_bilan.cell(row=r, column=c).fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    ws_bilan.cell(row=r, column=c).font = Font(name='Calibri', bold=True, size=11)
r += 2

ws_bilan.cell(row=r, column=1, value="* 2025 : exercice en cours, données au 20/07/2025 (non clôturé)")
ws_bilan.cell(row=r, column=1).font = Font(name='Calibri', italic=True, size=9)


# ============================================================
# FEUILLE 2 : COMPTE DE RÉSULTAT COMPARATIF
# ============================================================
ws_cr = wb.create_sheet("Compte de Résultat")
ws_cr.column_dimensions['A'].width = 50
ws_cr.column_dimensions['B'].width = 18
ws_cr.column_dimensions['C'].width = 18
ws_cr.column_dimensions['D'].width = 18

r = 1
ws_cr.merge_cells('A1:D1')
cell = ws_cr.cell(row=r, column=1, value="COMPTE DE RÉSULTAT COMPARATIF PAR NATURE")
cell.font = title_font
cell.alignment = Alignment(horizontal='center')
r += 1
ws_cr.merge_cells('A2:D2')
cell = ws_cr.cell(row=r, column=1, value="SDC 17 Rue Léon Cladel, Sèvres - Copropriété PERDUE")
cell.font = Font(name='Calibri', italic=True, size=11)
cell.alignment = Alignment(horizontal='center')
r += 2

ws_cr.cell(row=r, column=1, value="CHARGES PAR NATURE")
ws_cr.cell(row=r, column=2, value="2023\n(01/01 - 31/12)")
ws_cr.cell(row=r, column=3, value="2024\n(01/01 - 31/12)")
ws_cr.cell(row=r, column=4, value="2025\n(01/01 - 20/07)*")
style_header_row(ws_cr, r, 4)
r += 1

# --- SOUS-CLASSE 60 : Achats / Fluides ---
style_section_row(ws_cr, r, 4)
ws_cr.cell(row=r, column=1, value="60 - ACHATS ET FLUIDES")
r += 1
write_data_row(ws_cr, r, "6010 - Facture d'eau", 1029.59, 1018.55, 267.62, indent=1)
r += 1
write_data_row(ws_cr, r, "6020 - Consommation Électricité", 4212.02, 427.04, 2.35, indent=1)
r += 1
write_data_row(ws_cr, r, "Sous-total Achats / Fluides", 5241.61, 1445.59, 269.97, is_total=True)
for c in range(1, 5):
    ws_cr.cell(row=r, column=c).fill = subtotal_fill
r += 2

# --- SOUS-CLASSE 61 : Services extérieurs ---
style_section_row(ws_cr, r, 4)
ws_cr.cell(row=r, column=1, value="61 - SERVICES EXTÉRIEURS")
r += 1
write_data_row(ws_cr, r, "6110 - Entretien nettoyage", 1864.50, 0, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "6140 0015 - Entretien divers", 0, 260.70, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "6140 0018 - Contrat entretien jardins & espaces verts", 0, 1167.00, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "6140 0019 - Contrat service informatique (archives)", 26.51, 28.24, 29.75, indent=1)
r += 1
write_data_row(ws_cr, r, "6140 0023 - Contrat sécurité incendie", 346.06, 362.35, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "6150 0030 - Travaux divers", 1411.66, 1281.50, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "6160 0001 - Assurance multi-risques", 939.82, 948.46, 998.97, indent=1)
r += 1
write_data_row(ws_cr, r, "6160 0003 - Protection juridique", 138.00, 138.00, 140.48, indent=1)
r += 1
st61_2023 = 1864.50 + 26.51 + 346.06 + 1411.66 + 939.82 + 138.00
st61_2024 = 260.70 + 1167.00 + 28.24 + 362.35 + 1281.50 + 948.46 + 138.00
st61_2025 = 29.75 + 998.97 + 140.48
write_data_row(ws_cr, r, "Sous-total Services extérieurs", st61_2023, st61_2024, st61_2025, is_total=True)
for c in range(1, 5):
    ws_cr.cell(row=r, column=c).fill = subtotal_fill
r += 2

# --- SOUS-CLASSE 62 : Autres services extérieurs ---
style_section_row(ws_cr, r, 4)
ws_cr.cell(row=r, column=1, value="62 - AUTRES SERVICES EXTÉRIEURS")
r += 1
write_data_row(ws_cr, r, "6211 - Honoraires gestion syndic", 1800.00, 1800.00, 900.00, indent=1)
r += 1
write_data_row(ws_cr, r, "6213 0001 - Frais d'acheminement", 0, 10.99, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "6213 0002 - Affranchissements", 82.11, 62.41, 26.38, indent=1)
r += 1
write_data_row(ws_cr, r, "6222 0001 - Vacations complémentaires", 0, 96.00, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "6222 0002 - Réunions CS & AG complémentaires", 240.00, 0, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "6222 0005 - Pilotage et suivi contentieux", 1668.00, 0, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "6230 0002 - Honoraires huissiers", 495.40, 0, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "6230 0003 - Honoraires avocats", 4500.08, 0, 0, indent=1)
r += 1
st62_2023 = 1800 + 82.11 + 240 + 1668 + 495.40 + 4500.08
st62_2024 = 1800 + 10.99 + 62.41 + 96
st62_2025 = 900 + 26.38
write_data_row(ws_cr, r, "Sous-total Autres services extérieurs", st62_2023, st62_2024, st62_2025, is_total=True)
for c in range(1, 5):
    ws_cr.cell(row=r, column=c).fill = subtotal_fill
r += 2

# --- SOUS-CLASSE 66 : Charges financières ---
style_section_row(ws_cr, r, 4)
ws_cr.cell(row=r, column=1, value="66 - CHARGES FINANCIÈRES")
r += 1
write_data_row(ws_cr, r, "6620 0001 - Frais bancaires", 286.54, 96.00, 72.00, indent=1)
r += 1
write_data_row(ws_cr, r, "6620 0002 - Reliquat de répartition (rompus)", 0.25, 0.06, 0, indent=1)
r += 1
st66_2023 = 286.54 + 0.25
st66_2024 = 96.00 + 0.06
st66_2025 = 72.00
write_data_row(ws_cr, r, "Sous-total Charges financières", st66_2023, st66_2024, st66_2025, is_total=True)
for c in range(1, 5):
    ws_cr.cell(row=r, column=c).fill = subtotal_fill
r += 2

# --- SOUS-CLASSE 67 : Charges exceptionnelles ---
style_section_row(ws_cr, r, 4)
ws_cr.cell(row=r, column=1, value="67 - CHARGES EXCEPTIONNELLES")
r += 1
write_data_row(ws_cr, r, "6780 0003 - Charges except. sur exercices antérieurs", 25.49, 0, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "Sous-total Charges exceptionnelles", 25.49, 0, 0, is_total=True)
for c in range(1, 5):
    ws_cr.cell(row=r, column=c).fill = subtotal_fill
r += 2

# TOTAL CHARGES COURANTES
total_charges_2023 = 5241.61 + st61_2023 + st62_2023 + st66_2023 + 25.49
total_charges_2024 = 1445.59 + st61_2024 + st62_2024 + st66_2024
total_charges_2025 = 269.97 + st61_2025 + st62_2025 + st66_2025
write_data_row(ws_cr, r, "TOTAL CHARGES COURANTES", total_charges_2023, total_charges_2024, total_charges_2025, is_total=True)
for c in range(1, 5):
    ws_cr.cell(row=r, column=c).fill = PatternFill(start_color='F4B084', end_color='F4B084', fill_type='solid')
    ws_cr.cell(row=r, column=c).font = Font(name='Calibri', bold=True, size=11)
r += 3

# --- CHARGES TRAVAUX ---
style_section_row(ws_cr, r, 4)
ws_cr.cell(row=r, column=1, value="CHARGES TRAVAUX (hors budget courant)")
r += 1
write_data_row(ws_cr, r, "6221 0201 - Honoraires travaux (AG Réso 22 haie)", 120.00, 120.00, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "6710 0201 - Travaux décidés par l'AG (VMC/conduits)", 0, 792.00, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "TOTAL CHARGES TRAVAUX", 120.00, 912.00, 0, is_total=True)
for c in range(1, 5):
    ws_cr.cell(row=r, column=c).fill = PatternFill(start_color='F4B084', end_color='F4B084', fill_type='solid')
r += 3

# === PRODUITS ===
ws_cr.cell(row=r, column=1, value="PRODUITS PAR NATURE")
ws_cr.cell(row=r, column=2, value="2023")
ws_cr.cell(row=r, column=3, value="2024")
ws_cr.cell(row=r, column=4, value="2025*")
style_header_row(ws_cr, r, 4)
r += 1

style_section_row(ws_cr, r, 4)
ws_cr.cell(row=r, column=1, value="70 - PROVISIONS POUR CHARGES")
r += 1
write_data_row(ws_cr, r, "7010 - Provisions sur opérations courantes", 18000.00, 16000.00, 10274.94, indent=1)
r += 1
write_data_row(ws_cr, r, "Sous-total Provisions", 18000.00, 16000.00, 10274.94, is_total=True)
for c in range(1, 5):
    ws_cr.cell(row=r, column=c).fill = subtotal_fill
r += 2

style_section_row(ws_cr, r, 4)
ws_cr.cell(row=r, column=1, value="71 - PRODUITS EXCEPTIONNELS")
r += 1
write_data_row(ws_cr, r, "7180 - Produits exceptionnels", 400.00, 0, 0, indent=1)
r += 1
write_data_row(ws_cr, r, "Sous-total Produits exceptionnels", 400.00, 0, 0, is_total=True)
for c in range(1, 5):
    ws_cr.cell(row=r, column=c).fill = subtotal_fill
r += 2

total_produits_2023 = 18400.00
total_produits_2024 = 16000.00
total_produits_2025 = 10274.94

write_data_row(ws_cr, r, "TOTAL PRODUITS", total_produits_2023, total_produits_2024, total_produits_2025, is_total=True)
for c in range(1, 5):
    ws_cr.cell(row=r, column=c).fill = PatternFill(start_color='A9D18E', end_color='A9D18E', fill_type='solid')
    ws_cr.cell(row=r, column=c).font = Font(name='Calibri', bold=True, size=11)
r += 3

# === RÉSULTAT ===
ws_cr.cell(row=r, column=1, value="RÉSULTAT DE L'EXERCICE")
ws_cr.cell(row=r, column=2, value="2023")
ws_cr.cell(row=r, column=3, value="2024")
ws_cr.cell(row=r, column=4, value="2025*")
style_header_row(ws_cr, r, 4)
r += 1

res_courant_2023 = total_produits_2023 - total_charges_2023
res_courant_2024 = total_produits_2024 - total_charges_2024
res_courant_2025 = total_produits_2025 - total_charges_2025

write_data_row(ws_cr, r,
    "Résultat courant (Produits - Charges courantes)",
    res_courant_2023, res_courant_2024, res_courant_2025, is_total=True)
for c in range(1, 5):
    color = 'C6EFCE' if res_courant_2023 >= 0 else 'FFC7CE'
    ws_cr.cell(row=r, column=c).fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    ws_cr.cell(row=r, column=c).font = Font(name='Calibri', bold=True, size=11)
r += 1

write_data_row(ws_cr, r, "Résultat = Excédent (+) / Déficit (-)",
    res_courant_2023, res_courant_2024, res_courant_2025, is_total=True)
r += 1

ws_cr.cell(row=r, column=1, value="Un résultat positif = excédent (provisions > charges).")
ws_cr.cell(row=r, column=1).font = Font(name='Calibri', italic=True, size=9)
r += 1
ws_cr.cell(row=r, column=1, value="Un résultat négatif = déficit (charges > provisions).")
ws_cr.cell(row=r, column=1).font = Font(name='Calibri', italic=True, size=9)
r += 2
ws_cr.cell(row=r, column=1, value="* 2025 : exercice en cours au 20/07/2025, données partielles (6 mois et demi)")
ws_cr.cell(row=r, column=1).font = Font(name='Calibri', italic=True, size=9, color='FF0000')
r += 1
ws_cr.cell(row=r, column=1, value="Note : les charges travaux (6221/6710) sont hors budget courant et font l'objet d'appels spécifiques.")
ws_cr.cell(row=r, column=1).font = Font(name='Calibri', italic=True, size=9)


# ============================================================
# FEUILLE 3 : DÉTAIL BILAN PAR COMPTE
# ============================================================
ws_detail = wb.create_sheet("Détail Bilan par Compte")
ws_detail.column_dimensions['A'].width = 12
ws_detail.column_dimensions['B'].width = 40
ws_detail.column_dimensions['C'].width = 18
ws_detail.column_dimensions['D'].width = 18
ws_detail.column_dimensions['E'].width = 18

r = 1
ws_detail.merge_cells('A1:E1')
cell = ws_detail.cell(row=r, column=1, value="DÉTAIL DU BILAN PAR COMPTE")
cell.font = title_font
cell.alignment = Alignment(horizontal='center')
r += 2

for c_idx, val in enumerate(["Compte", "Libellé", "2023", "2024", "2025*"], 1):
    cell = ws_detail.cell(row=r, column=c_idx, value=val)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center')
    cell.border = thin_border
r += 1

detail_data = [
    ("", "ACTIF", None, None, None, True),
    ("1200", "Solde en attente sur travaux", 120.00, 0, 0, False),
    ("4010", "Fournisseurs débiteurs", 0, 0, 278.48, False),
    ("450", "Copropriétaires débiteurs", 0, 0, 3776.31, False),
    ("4711", "Régularisation charges - Débiteur", 18665.97, 26363.21, 26363.21, False),
    ("4712", "Régularisation travaux - Débiteur", 0, 911.97, 911.97, False),
    ("4730", "Rompus débiteurs", 0.06, 0.09, 0.09, False),
    ("5120 0001", "Compte courant Montepaschi", 2968.03, 9525.90, 13074.46, False),
    ("5120 0004", "Livret A Monte Paschi", 910.22, 1742.55, 1967.53, False),
    ("", "TOTAL ACTIF", 22664.28, 38543.72, 46372.05, True),
    ("", "", None, None, None, False),
    ("", "PASSIF", None, None, None, True),
    ("1031", "Avances de trésorerie", 2000.00, 2000.00, 2000.00, False),
    ("1050", "Fonds de travaux", 910.22, 1742.55, 2256.18, False),
    ("4010", "Fournisseurs créditeurs", 1330.06, 659.45, 0, False),
    ("4080", "Factures non parvenues", 0, 141.72, 0, False),
    ("4630", "Tiers créditeurs (anciens copro)", 424.00, 0, 0, False),
    ("4721", "Régularisation charges - Créditeur", 18000.00, 34000.00, 34000.00, False),
    ("", "Résultat en cours (classe 6 & 7)", 0, 0, 8115.87, False),
    ("", "TOTAL PASSIF", 22664.28, 38543.72, 46372.05, True),
]

for row_data in detail_data:
    compte, libelle, v23, v24, v25, bold = row_data
    ws_detail.cell(row=r, column=1, value=compte).font = total_font if bold else normal_font
    ws_detail.cell(row=r, column=1).border = thin_border
    ws_detail.cell(row=r, column=2, value=libelle).font = total_font if bold else normal_font
    ws_detail.cell(row=r, column=2).border = thin_border
    for ci, v in enumerate([v23, v24, v25], 3):
        cell = ws_detail.cell(row=r, column=ci, value=v)
        cell.number_format = euro_format if v is not None else ''
        cell.font = total_font if bold else normal_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='right')
    if bold:
        for c in range(1, 6):
            ws_detail.cell(row=r, column=c).fill = total_fill
    r += 1


# --- SAVE ---
output_path = "/workspace/bilan_comptable_comparatif.xlsx"
wb.save(output_path)
print(f"Fichier généré : {output_path}")
print(f"Résultat 2023 : {res_courant_2023:,.2f} €")
print(f"Résultat 2024 : {res_courant_2024:,.2f} €")
print(f"Résultat 2025 (partiel) : {res_courant_2025:,.2f} €")
