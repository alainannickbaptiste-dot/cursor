import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

wb = openpyxl.Workbook()

# --- STYLES ---
title_font = Font(name='Calibri', bold=True, size=14)
header_font = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
section_font = Font(name='Calibri', bold=True, size=11)
normal_font = Font(name='Calibri', size=10)
total_font = Font(name='Calibri', bold=True, size=10)

header_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
section_fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
total_fill = PatternFill(start_color='B4C6E7', end_color='B4C6E7', fill_type='solid')
subtotal_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
result_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
charges_total_fill = PatternFill(start_color='F4B084', end_color='F4B084', fill_type='solid')
produits_total_fill = PatternFill(start_color='A9D18E', end_color='A9D18E', fill_type='solid')

thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

euro_format = '#,##0.00 €'


def style_header_row(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin_border


def style_section_row(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = section_font
        cell.fill = section_fill
        cell.border = thin_border


def write_row(ws, row, label, v2023, v2024, v2025, indent=0, bold=False, fill=None):
    cell_a = ws.cell(row=row, column=1, value=(' ' * indent * 3) + label)
    cell_a.font = total_font if bold else normal_font
    cell_a.border = thin_border
    cell_a.alignment = Alignment(wrap_text=True)
    if fill:
        cell_a.fill = fill

    for col_idx, val in enumerate([v2023, v2024, v2025], start=2):
        cell = ws.cell(row=row, column=col_idx, value=val)
        if val is not None:
            cell.number_format = euro_format
        cell.font = total_font if bold else normal_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='right')
        if fill:
            cell.fill = fill


# ============================================================
# DONNÉES VÉRIFIÉES - Comptes de charges (classe 6)
# ============================================================
# Montants nets transférés au décompte, vérifiés contre 4711/4721

# 2023 - Charges (vérifiées : total = 19 186,03 ; 4711=18 665,97 + rompus 0,06 + 7180 produit 400)
c6_2023 = {
    '6010': ('Facture d\'eau', 1029.59),
    '6020': ('Consommation Électricité', 4212.02),
    '6110': ('Entretien nettoyage', 1864.50),
    '6140_19': ('Contrat service informatique (archives)', 26.51),
    '6140_23': ('Contrat de sécurité incendie', 346.06),
    '6150': ('Travaux divers', 1411.66),
    '6160_01': ('Assurance multi-risques', 939.82),
    '6160_03': ('Protection juridique', 138.00),
    '6211': ('Honoraires gestion syndic', 1800.00),
    '6213_02': ('Affranchissements', 82.11),
    '6221': ('Honoraires travaux (AG Réso 22 haie)', 120.00),
    '6222_02': ('Réunions CS & AG complémentaires', 240.00),
    '6222_05': ('Pilotage et suivi contentieux', 1668.00),
    '6230_02': ('Honoraires huissiers', 495.40),
    '6230_03': ('Honoraires avocats', 4500.08),
    '6620_01': ('Frais bancaires', 286.54),
    '6620_02': ('Reliquat de répartition (rompus)', 0.25),
    '6780': ('Charges except. sur exercices antérieurs', 25.49),
}

# 2024 - Charges (vérifiées : total = 8 609,30 ; 4711=7 697,24 + rompus 0,06 + travaux 912,00)
c6_2024 = {
    '6010': ('Facture d\'eau', 1018.55),
    '6020': ('Consommation Électricité', 427.04),
    '6140_15': ('Entretien divers', 260.70),
    '6140_18': ('Contrat entretien jardins & espaces verts', 1167.00),
    '6140_19': ('Contrat service informatique (archives)', 28.24),
    '6140_23': ('Contrat de sécurité incendie', 362.35),
    '6150': ('Travaux divers (disjoncteur)', 1281.50),
    '6160_01': ('Assurance multi-risques', 948.46),
    '6160_03': ('Protection juridique', 138.00),
    '6211': ('Honoraires gestion syndic', 1800.00),
    '6213_01': ('Frais d\'acheminement', 10.99),
    '6213_02': ('Affranchissements', 62.41),
    '6221': ('Honoraires travaux (AG Réso 22 haie)', 120.00),
    '6222_01': ('Vacations complémentaires', 96.00),
    '6620_01': ('Frais bancaires', 96.00),
    '6620_02': ('Reliquat de répartition (rompus)', 0.06),
    '6710': ('Travaux décidés par l\'AG (VMC/conduits)', 792.00),
}

# 2025 - Charges partielles au 20/07/2025 (vérifiées : total classe 6 net = 2 159,07)
c6_2025 = {
    '6010': ('Facture d\'eau', 267.62),
    '6020': ('Consommation Électricité', 2.35),
    '6140_19': ('Contrat service informatique (archives)', 29.75),
    '6160_01': ('Assurance multi-risques', 998.97),
    '6160_03': ('Protection juridique', -138.00),
    '6211': ('Honoraires gestion syndic', 900.00),
    '6213_02': ('Affranchissements', 26.38),
    '6620_01': ('Frais bancaires', 72.00),
}

# Produits (classe 7)
c7_2023 = {
    '7010': ('Provisions sur opérations courantes', 18000.00),
    '7180': ('Produits exceptionnels', 400.00),
}
c7_2024 = {
    '7010': ('Provisions sur opérations courantes', 16000.00),
}
c7_2025 = {
    '7010': ('Provisions sur opérations courantes', 10274.94),
}

total_c6_2023 = sum(v for _, v in c6_2023.values())
total_c6_2024 = sum(v for _, v in c6_2024.values())
total_c6_2025 = sum(v for _, v in c6_2025.values())
total_c7_2023 = sum(v for _, v in c7_2023.values())
total_c7_2024 = sum(v for _, v in c7_2024.values())
total_c7_2025 = sum(v for _, v in c7_2025.values())

res_2023 = total_c7_2023 - total_c6_2023
res_2024 = total_c7_2024 - total_c6_2024
res_2025 = total_c7_2025 - total_c6_2025


# ============================================================
# FEUILLE 1 : BILAN COMPARATIF
# ============================================================
ws = wb.active
ws.title = "Bilan Comparatif"
ws.column_dimensions['A'].width = 48
ws.column_dimensions['B'].width = 18
ws.column_dimensions['C'].width = 18
ws.column_dimensions['D'].width = 18

r = 1
ws.merge_cells('A1:D1')
ws.cell(row=r, column=1, value="BILAN COMPARATIF - SDC 17 Rue Léon Cladel, Sèvres").font = title_font
ws.cell(row=r, column=1).alignment = Alignment(horizontal='center')
r += 1
ws.merge_cells('A2:D2')
ws.cell(row=r, column=1, value="Copropriété PERDUE - Syndic MANDA").font = Font(name='Calibri', italic=True, size=11)
ws.cell(row=r, column=1).alignment = Alignment(horizontal='center')
r += 2

ws.cell(row=r, column=1, value="ACTIF")
ws.cell(row=r, column=2, value="31/12/2023")
ws.cell(row=r, column=3, value="31/12/2024")
ws.cell(row=r, column=4, value="20/07/2025*")
style_header_row(ws, r, 4)
r += 1

# ACTIF IMMOBILISÉ
style_section_row(ws, r, 4)
ws.cell(row=r, column=1, value="ACTIF IMMOBILISÉ")
r += 1
write_row(ws, r, "Solde en attente sur travaux (1200)", 120.00, 0, 0, indent=1)
r += 1
write_row(ws, r, "Total Actif immobilisé", 120.00, 0, 0, bold=True, fill=total_fill)
r += 2

# ACTIF CIRCULANT
style_section_row(ws, r, 4)
ws.cell(row=r, column=1, value="ACTIF CIRCULANT")
r += 1
write_row(ws, r, "Fournisseurs - soldes débiteurs (4010)", 0, 0, 278.48, indent=1)
r += 1
write_row(ws, r, "Copropriétaires - soldes débiteurs (450)", 0, 0, 3776.31, indent=1)
r += 1
write_row(ws, r, "Régularisation charges - Débiteur (4711)", 18665.97, 26363.21, 26363.21, indent=1)
r += 1
write_row(ws, r, "Régularisation travaux - Débiteur (4712)", 0, 911.97, 911.97, indent=1)
r += 1
write_row(ws, r, "Rompus débiteurs (4730)", 0.06, 0.09, 0.09, indent=1)
r += 1
write_row(ws, r, "Total Actif circulant", 18666.03, 27275.27, 31330.06, bold=True, fill=total_fill)
r += 2

# TRÉSORERIE
style_section_row(ws, r, 4)
ws.cell(row=r, column=1, value="TRÉSORERIE ACTIVE")
r += 1
write_row(ws, r, "Compte courant Montepaschi (5120 0001)", 2968.03, 9525.90, 13074.46, indent=1)
r += 1
write_row(ws, r, "Livret A Monte Paschi (5120 0004)", 910.22, 1742.55, 1967.53, indent=1)
r += 1
write_row(ws, r, "Total Trésorerie", 3878.25, 11268.45, 15041.99, bold=True, fill=total_fill)
r += 2

write_row(ws, r, "TOTAL ACTIF", 22664.28, 38543.72, 46372.05, bold=True, fill=result_fill)
for c in range(1, 5):
    ws.cell(row=r, column=c).font = Font(name='Calibri', bold=True, size=11)
r += 3

# PASSIF
ws.cell(row=r, column=1, value="PASSIF")
ws.cell(row=r, column=2, value="31/12/2023")
ws.cell(row=r, column=3, value="31/12/2024")
ws.cell(row=r, column=4, value="20/07/2025*")
style_header_row(ws, r, 4)
r += 1

style_section_row(ws, r, 4)
ws.cell(row=r, column=1, value="CAPITAUX PROPRES & AVANCES")
r += 1
write_row(ws, r, "Avances de trésorerie (1031)", 2000.00, 2000.00, 2000.00, indent=1)
r += 1
write_row(ws, r, "Fonds de travaux (1050)", 910.22, 1742.55, 2256.18, indent=1)
r += 1
write_row(ws, r, "Total Capitaux propres & Avances", 2910.22, 3742.55, 4256.18, bold=True, fill=total_fill)
r += 2

style_section_row(ws, r, 4)
ws.cell(row=r, column=1, value="DETTES")
r += 1
write_row(ws, r, "Fournisseurs créditeurs (4010)", 1330.06, 659.45, 0, indent=1)
r += 1
write_row(ws, r, "Factures non parvenues (4080)", 0, 141.72, 0, indent=1)
r += 1
write_row(ws, r, "Tiers créditeurs - anciens copro (4630)", 424.00, 0, 0, indent=1)
r += 1
write_row(ws, r, "Régularisation charges - Créditeur (4721)", 18000.00, 34000.00, 34000.00, indent=1)
r += 1
write_row(ws, r, "Total Dettes", 19754.06, 34801.17, 34000.00, bold=True, fill=total_fill)
r += 2

style_section_row(ws, r, 4)
ws.cell(row=r, column=1, value="RÉSULTAT EN COURS (exercice 2025 non clôturé)")
r += 1
write_row(ws, r, "Provisions appelées - Charges engagées", 0, 0, 8115.87, indent=1)
r += 1
write_row(ws, r, "Total Résultat en cours", 0, 0, 8115.87, bold=True, fill=total_fill)
r += 2

write_row(ws, r, "TOTAL PASSIF", 22664.28, 38543.72, 46372.05, bold=True, fill=result_fill)
for c in range(1, 5):
    ws.cell(row=r, column=c).font = Font(name='Calibri', bold=True, size=11)
r += 2
ws.cell(row=r, column=1, value="* 2025 : exercice en cours au 20/07/2025 (non clôturé)").font = Font(name='Calibri', italic=True, size=9)


# ============================================================
# FEUILLE 2 : COMPTE DE RÉSULTAT COMPARATIF PAR NATURE
# ============================================================
ws2 = wb.create_sheet("Compte de Résultat")
ws2.column_dimensions['A'].width = 55
ws2.column_dimensions['B'].width = 18
ws2.column_dimensions['C'].width = 18
ws2.column_dimensions['D'].width = 18

r = 1
ws2.merge_cells('A1:D1')
ws2.cell(row=r, column=1, value="COMPTE DE RÉSULTAT COMPARATIF PAR NATURE DE COMPTE").font = title_font
ws2.cell(row=r, column=1).alignment = Alignment(horizontal='center')
r += 1
ws2.merge_cells('A2:D2')
ws2.cell(row=r, column=1, value="SDC 17 Rue Léon Cladel, Sèvres - Copropriété PERDUE").font = Font(name='Calibri', italic=True, size=11)
ws2.cell(row=r, column=1).alignment = Alignment(horizontal='center')
r += 2

# --- CHARGES ---
ws2.cell(row=r, column=1, value="CHARGES (Classe 6)")
ws2.cell(row=r, column=2, value="2023\n(01/01 - 31/12)")
ws2.cell(row=r, column=3, value="2024\n(01/01 - 31/12)")
ws2.cell(row=r, column=4, value="2025\n(01/01 - 20/07)*")
style_header_row(ws2, r, 4)
r += 1

def val(accounts, key):
    return accounts.get(key, (None, 0))[1] if key in accounts else 0

# Sous-classe 60
style_section_row(ws2, r, 4)
ws2.cell(row=r, column=1, value="60 - ACHATS ET FLUIDES")
r += 1
write_row(ws2, r, "6010 - Facture d'eau", val(c6_2023, '6010'), val(c6_2024, '6010'), val(c6_2025, '6010'), indent=1)
r += 1
write_row(ws2, r, "6020 - Consommation Électricité", val(c6_2023, '6020'), val(c6_2024, '6020'), val(c6_2025, '6020'), indent=1)
r += 1
st60 = [val(c6_2023,'6010')+val(c6_2023,'6020'), val(c6_2024,'6010')+val(c6_2024,'6020'), val(c6_2025,'6010')+val(c6_2025,'6020')]
write_row(ws2, r, "Sous-total 60 - Achats / Fluides", st60[0], st60[1], st60[2], bold=True, fill=subtotal_fill)
r += 2

# Sous-classe 61
style_section_row(ws2, r, 4)
ws2.cell(row=r, column=1, value="61 - SERVICES EXTÉRIEURS")
r += 1
write_row(ws2, r, "6110 - Entretien nettoyage", val(c6_2023,'6110'), 0, 0, indent=1)
r += 1
write_row(ws2, r, "6140 0015 - Entretien divers", 0, val(c6_2024,'6140_15'), 0, indent=1)
r += 1
write_row(ws2, r, "6140 0018 - Contrat entretien jardins & espaces verts", 0, val(c6_2024,'6140_18'), 0, indent=1)
r += 1
write_row(ws2, r, "6140 0019 - Contrat service informatique (archives)", val(c6_2023,'6140_19'), val(c6_2024,'6140_19'), val(c6_2025,'6140_19'), indent=1)
r += 1
write_row(ws2, r, "6140 0023 - Contrat de sécurité incendie", val(c6_2023,'6140_23'), val(c6_2024,'6140_23'), 0, indent=1)
r += 1
write_row(ws2, r, "6150 0030 - Travaux divers", val(c6_2023,'6150'), val(c6_2024,'6150'), 0, indent=1)
r += 1
write_row(ws2, r, "6160 0001 - Assurance multi-risques", val(c6_2023,'6160_01'), val(c6_2024,'6160_01'), val(c6_2025,'6160_01'), indent=1)
r += 1
write_row(ws2, r, "6160 0003 - Protection juridique", val(c6_2023,'6160_03'), val(c6_2024,'6160_03'), val(c6_2025,'6160_03'), indent=1)
r += 1
st61 = [
    val(c6_2023,'6110')+val(c6_2023,'6140_19')+val(c6_2023,'6140_23')+val(c6_2023,'6150')+val(c6_2023,'6160_01')+val(c6_2023,'6160_03'),
    val(c6_2024,'6140_15')+val(c6_2024,'6140_18')+val(c6_2024,'6140_19')+val(c6_2024,'6140_23')+val(c6_2024,'6150')+val(c6_2024,'6160_01')+val(c6_2024,'6160_03'),
    val(c6_2025,'6140_19')+val(c6_2025,'6160_01')+val(c6_2025,'6160_03'),
]
write_row(ws2, r, "Sous-total 61 - Services extérieurs", st61[0], st61[1], st61[2], bold=True, fill=subtotal_fill)
r += 2

# Sous-classe 62
style_section_row(ws2, r, 4)
ws2.cell(row=r, column=1, value="62 - AUTRES SERVICES EXTÉRIEURS (honoraires, frais juridiques)")
r += 1
write_row(ws2, r, "6211 - Honoraires gestion syndic", val(c6_2023,'6211'), val(c6_2024,'6211'), val(c6_2025,'6211'), indent=1)
r += 1
write_row(ws2, r, "6213 0001 - Frais d'acheminement", 0, val(c6_2024,'6213_01'), 0, indent=1)
r += 1
write_row(ws2, r, "6213 0002 - Affranchissements", val(c6_2023,'6213_02'), val(c6_2024,'6213_02'), val(c6_2025,'6213_02'), indent=1)
r += 1
write_row(ws2, r, "6221 0201 - Honoraires travaux (AG Réso 22 haie)", val(c6_2023,'6221'), val(c6_2024,'6221'), 0, indent=1)
r += 1
write_row(ws2, r, "6222 0001 - Vacations complémentaires", 0, val(c6_2024,'6222_01'), 0, indent=1)
r += 1
write_row(ws2, r, "6222 0002 - Réunions CS & AG complémentaires", val(c6_2023,'6222_02'), 0, 0, indent=1)
r += 1
write_row(ws2, r, "6222 0005 - Pilotage et suivi contentieux", val(c6_2023,'6222_05'), 0, 0, indent=1)
r += 1
write_row(ws2, r, "6230 0002 - Honoraires huissiers", val(c6_2023,'6230_02'), 0, 0, indent=1)
r += 1
write_row(ws2, r, "6230 0003 - Honoraires avocats", val(c6_2023,'6230_03'), 0, 0, indent=1)
r += 1
st62 = [
    val(c6_2023,'6211')+val(c6_2023,'6213_02')+val(c6_2023,'6221')+val(c6_2023,'6222_02')+val(c6_2023,'6222_05')+val(c6_2023,'6230_02')+val(c6_2023,'6230_03'),
    val(c6_2024,'6211')+val(c6_2024,'6213_01')+val(c6_2024,'6213_02')+val(c6_2024,'6221')+val(c6_2024,'6222_01'),
    val(c6_2025,'6211')+val(c6_2025,'6213_02'),
]
write_row(ws2, r, "Sous-total 62 - Autres services extérieurs", st62[0], st62[1], st62[2], bold=True, fill=subtotal_fill)
r += 2

# Sous-classe 66
style_section_row(ws2, r, 4)
ws2.cell(row=r, column=1, value="66 - CHARGES FINANCIÈRES")
r += 1
write_row(ws2, r, "6620 0001 - Frais bancaires", val(c6_2023,'6620_01'), val(c6_2024,'6620_01'), val(c6_2025,'6620_01'), indent=1)
r += 1
write_row(ws2, r, "6620 0002 - Reliquat de répartition (rompus)", val(c6_2023,'6620_02'), val(c6_2024,'6620_02'), 0, indent=1)
r += 1
st66 = [
    val(c6_2023,'6620_01')+val(c6_2023,'6620_02'),
    val(c6_2024,'6620_01')+val(c6_2024,'6620_02'),
    val(c6_2025,'6620_01'),
]
write_row(ws2, r, "Sous-total 66 - Charges financières", st66[0], st66[1], st66[2], bold=True, fill=subtotal_fill)
r += 2

# Sous-classe 67
style_section_row(ws2, r, 4)
ws2.cell(row=r, column=1, value="67 - CHARGES EXCEPTIONNELLES / TRAVAUX AG")
r += 1
write_row(ws2, r, "6780 0003 - Charges except. sur exercices antérieurs", val(c6_2023,'6780'), 0, 0, indent=1)
r += 1
write_row(ws2, r, "6710 0201 - Travaux décidés par l'AG (VMC/conduits)", 0, val(c6_2024,'6710'), 0, indent=1)
r += 1
st67 = [val(c6_2023,'6780'), val(c6_2024,'6710'), 0]
write_row(ws2, r, "Sous-total 67 - Charges exceptionnelles / Travaux", st67[0], st67[1], st67[2], bold=True, fill=subtotal_fill)
r += 2

# TOTAL CHARGES
write_row(ws2, r, "TOTAL CHARGES (Classe 6)", total_c6_2023, total_c6_2024, total_c6_2025, bold=True, fill=charges_total_fill)
for c in range(1, 5):
    ws2.cell(row=r, column=c).font = Font(name='Calibri', bold=True, size=11)
r += 3

# --- PRODUITS ---
ws2.cell(row=r, column=1, value="PRODUITS (Classe 7)")
ws2.cell(row=r, column=2, value="2023")
ws2.cell(row=r, column=3, value="2024")
ws2.cell(row=r, column=4, value="2025*")
style_header_row(ws2, r, 4)
r += 1

style_section_row(ws2, r, 4)
ws2.cell(row=r, column=1, value="70 - PROVISIONS POUR CHARGES COURANTES")
r += 1
write_row(ws2, r, "7010 - Provisions sur opérations courantes", 18000.00, 16000.00, 10274.94, indent=1)
r += 1
write_row(ws2, r, "Sous-total 70 - Provisions", 18000.00, 16000.00, 10274.94, bold=True, fill=subtotal_fill)
r += 2

style_section_row(ws2, r, 4)
ws2.cell(row=r, column=1, value="71 - PRODUITS EXCEPTIONNELS")
r += 1
write_row(ws2, r, "7180 - Produits exceptionnels (FNP antérieures)", 400.00, 0, 0, indent=1)
r += 1
write_row(ws2, r, "Sous-total 71 - Produits exceptionnels", 400.00, 0, 0, bold=True, fill=subtotal_fill)
r += 2

write_row(ws2, r, "TOTAL PRODUITS (Classe 7)", total_c7_2023, total_c7_2024, total_c7_2025, bold=True, fill=produits_total_fill)
for c in range(1, 5):
    ws2.cell(row=r, column=c).font = Font(name='Calibri', bold=True, size=11)
r += 3

# RÉSULTAT
ws2.cell(row=r, column=1, value="RÉSULTAT DE L'EXERCICE")
ws2.cell(row=r, column=2, value="2023")
ws2.cell(row=r, column=3, value="2024")
ws2.cell(row=r, column=4, value="2025*")
style_header_row(ws2, r, 4)
r += 1

write_row(ws2, r, "Résultat = Produits (Cl. 7) - Charges (Cl. 6)", res_2023, res_2024, res_2025, bold=True, fill=result_fill)
for c in range(1, 5):
    ws2.cell(row=r, column=c).font = Font(name='Calibri', bold=True, size=12)
r += 2

write_row(ws2, r, "Excédent (+) ou Déficit (-)", res_2023, res_2024, res_2025, bold=True)
r += 2

ws2.cell(row=r, column=1, value="Résultat positif = excédent (provisions > charges réelles).").font = Font(italic=True, size=9)
r += 1
ws2.cell(row=r, column=1, value="Résultat négatif = déficit (charges réelles > provisions appelées).").font = Font(italic=True, size=9)
r += 2
ws2.cell(row=r, column=1, value="* 2025 : exercice en cours au 20/07/2025, données partielles (~6,5 mois).").font = Font(italic=True, size=9, color='FF0000')
r += 1
ws2.cell(row=r, column=1, value="Note : les frais privatifs (mutations, mises en demeure) transitent en débit/crédit et n'affectent pas le résultat.").font = Font(italic=True, size=9)


# --- SAVE ---
output_path = "/workspace/bilan_comptable_comparatif.xlsx"
wb.save(output_path)

print(f"Fichier généré : {output_path}")
print()
print("=== VÉRIFICATION DES TOTAUX ===")
print(f"Total charges 2023 : {total_c6_2023:,.2f} €")
print(f"Total produits 2023 : {total_c7_2023:,.2f} €")
print(f"Résultat 2023 : {res_2023:,.2f} €")
print()
print(f"Total charges 2024 : {total_c6_2024:,.2f} €")
print(f"Total produits 2024 : {total_c7_2024:,.2f} €")
print(f"Résultat 2024 : {res_2024:,.2f} €")
print()
print(f"Total charges 2025 : {total_c6_2025:,.2f} €")
print(f"Total produits 2025 : {total_c7_2025:,.2f} €")
print(f"Résultat 2025 : {res_2025:,.2f} €")
print()
print("=== CONTRÔLE SOUS-TOTAUX 2023 ===")
print(f"  60: {st60[0]:,.2f}")
print(f"  61: {st61[0]:,.2f}")
print(f"  62: {st62[0]:,.2f}")
print(f"  66: {st66[0]:,.2f}")
print(f"  67: {st67[0]:,.2f}")
print(f"  Somme: {st60[0]+st61[0]+st62[0]+st66[0]+st67[0]:,.2f}")
