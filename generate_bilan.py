#!/usr/bin/env python3
"""
Parse les 3 grands livres PDF et produit un bilan comparatif et un compte
de résultat basé sur les MOUVEMENTS de l'exercice.
"""

import re
import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


# =====================================================================
# 1. EXTRACTION TEXTE DES PDF
# =====================================================================

def extract_text(pdf_path):
    """Extrait le texte de toutes les pages d'un PDF."""
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split('\n'):
                    lines.append(line)
    return lines


# =====================================================================
# 2. PARSER DU GRAND LIVRE
# =====================================================================

def parse_amount(s):
    """Parse un montant comme '2000.00' ou '295677.47' en float."""
    s = s.strip().replace(' ', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_grand_livre(pdf_path):
    """
    Parse un grand livre PDF et retourne un dict de comptes avec mouvements.

    Pour chaque compte, on accumule les débits et crédits des mouvements
    de l'exercice, en excluant les écritures techniques.

    On capture aussi le solde final (avant clôture) pour le bilan, via les
    colonnes "solde débiteur / solde créditeur" de la dernière écriture.

    Retourne: { compte_num: { 'label': str, 'debit': float, 'credit': float,
                               'solde_final_deb': float, 'solde_final_cred': float } }
    """
    lines = extract_text(pdf_path)
    accounts = {}
    current_account = None
    current_label = None

    # Pattern pour en-tête de compte (caractères doublés) : "11003311 00000011 ..."
    # Le numéro de compte normal est intercalé dans les doubles
    account_header_re = re.compile(
        r'^((\d)\2((\d)\4){3})\s+((\d)\6((\d)\8){3})\s+'
    )

    # Alternative: parfois le numéro est normal "450 0002" etc.
    account_header_simple_re = re.compile(r'^(\d{3,4})\s+(\d{4})\s+(.+)')

    # Pattern pour une ligne de mouvement : commence par dd/mm/yy
    move_date_re = re.compile(r'^(\d{2}/\d{2}/\d{2,4})\s+(.+)')

    # Pattern pour extraire les 4 montants à la fin d'une ligne
    amounts_re = re.compile(r'([\d.]+)€\s+([\d.]+)€\s+([\d.]+)€\s+([\d.]+)€\s*$')

    # Lignes à ignorer (mouvements techniques)
    skip_keywords = [
        'AA nnoouuvveeaauu',
        'A nouveau',
        'Cloture Exercice',
        'Cloture du compte',
        'TToottaall',
        'TTOOTTAAUUXX',
        'TOTAUX',
        'Total sous-classe',
        'Total classe',
        'ssoouuss--ccllaassssee',
        'ccllaassssee',
        'Décompte charges courantes',
        'Décompte travaux',
        'Charges et provisions travaux report',
        'Transfert avance',
        'Transfert fonds',
        'Transfert solde mutation',
        'Transfert provision',
        'Transfert compte',
        'report de charge pour la p',
        'GGrraanndd LLiivvrree',
        'Grand Livre',
        'Édité le',
        'DDAATTEE',
        'PPéérriiooddee',
        'CCoopprroopprriiééttéé',
        'VVoottrree SSyynnddiicc',
        '--- PAGE BREAK ---',
    ]

    def should_skip(text, account=None):
        for kw in skip_keywords:
            if kw in text:
                return True
        return False

    def decode_doubled(s):
        """Decode doubled characters: 'AAvvaanncceess' -> 'Avances'."""
        result = []
        i = 0
        while i < len(s):
            result.append(s[i])
            if i + 1 < len(s) and s[i] == s[i + 1]:
                i += 2
            else:
                i += 1
        return ''.join(result)

    def extract_account_num(doubled_str):
        """Extract account number from doubled format: '11003311' -> '1031'."""
        chars = []
        i = 0
        while i < len(doubled_str):
            chars.append(doubled_str[i])
            i += 2
        return ''.join(chars)

    i = 0
    while i < len(lines):
        line = lines[i].strip()

        # Skip empty lines
        if not line:
            i += 1
            continue

        # Detect account header with doubled characters
        m = account_header_re.match(line)
        if m:
            part1 = extract_account_num(m.group(1))
            part2 = extract_account_num(m.group(5))
            current_account = f"{part1} {part2}"
            # Label is the rest of the line, decoded
            rest = line[m.end():].strip()
            current_label = decode_doubled(rest).strip()
            # Remove duplicated label after tab
            if '\t' in current_label:
                current_label = current_label.split('\t')[0].strip()
            if current_account not in accounts:
                accounts[current_account] = {
                    'label': current_label,
                    'debit': 0.0,
                    'credit': 0.0,
                    'solde_final_deb': 0.0,
                    'solde_final_cred': 0.0,
                    'last_solde_deb': 0.0,
                    'last_solde_cred': 0.0,
                }
            i += 1
            continue

        # Detect movement line (starts with a date)
        dm = move_date_re.match(line)
        if dm and current_account:
            full_line = line
            # Check if amounts are on this line
            am = amounts_re.search(full_line)

            if not am:
                # Amounts might be on a continuation line
                j = i + 1
                while j < len(lines) and not amounts_re.search(full_line):
                    next_l = lines[j].strip()
                    if not next_l:
                        j += 1
                        continue
                    if any(kw in next_l for kw in ['GGrraanndd LLiivvrree', 'Grand Livre', 'Édité le', 'DDAATTEE', '--- PAGE BREAK ---']):
                        j += 1
                        continue
                    if move_date_re.match(next_l):
                        break
                    if account_header_re.match(next_l):
                        break
                    full_line += ' ' + next_l
                    j += 1
                am = amounts_re.search(full_line)
                if am:
                    i = j
                else:
                    i += 1
                continue

            debit = parse_amount(am.group(1))
            credit = parse_amount(am.group(2))
            solde_deb = parse_amount(am.group(3))
            solde_cred = parse_amount(am.group(4))

            is_cloture = 'Cloture' in full_line or 'comptes? d' in full_line

            # Always track the solde (before the cloture line resets it to 0)
            if not is_cloture:
                accounts[current_account]['last_solde_deb'] = solde_deb
                accounts[current_account]['last_solde_cred'] = solde_cred

            # Only add to movement totals if not a skip line
            if should_skip(full_line, current_account):
                i += 1
                continue

            accounts[current_account]['debit'] += debit
            accounts[current_account]['credit'] += credit

        i += 1

    # Store final soldes
    for k in accounts:
        accounts[k]['solde_final_deb'] = accounts[k]['last_solde_deb']
        accounts[k]['solde_final_cred'] = accounts[k]['last_solde_cred']

    return accounts


# =====================================================================
# 3. LIRE LES 3 FICHIERS
# =====================================================================

base = '/home/ubuntu/.cursor/projects/workspace/uploads'
print("Parsing 2023...")
acc_2023 = parse_grand_livre(f'{base}/grand_livre-du-2023-01-01-au-2023-12-31.pdf')
print("Parsing 2024...")
acc_2024 = parse_grand_livre(f'{base}/grand_livre-du-2024-01-01-au-2024-12-31.pdf')
print("Parsing 2025...")
acc_2025 = parse_grand_livre(f'{base}/grand_livre-du-2025-01-01-au-2025-12-31.pdf')


# =====================================================================
# 4. AFFICHAGE ET VÉRIFICATION
# =====================================================================

def get_class(k):
    return k.strip()[0]

def get_subclass(k):
    return k.strip()[:2]


def print_accounts(year, acc):
    print(f"\n=== Comptes {year} ===")
    for k in sorted(acc.keys()):
        d = acc[k]
        net = d['debit'] - d['credit']
        if abs(d['debit']) + abs(d['credit']) > 0.001:
            print(f"  {k:12s} {d['label'][:40]:40s}  D={d['debit']:>12,.2f}  C={d['credit']:>12,.2f}  Net={net:>12,.2f}")

print_accounts(2023, acc_2023)
print_accounts(2024, acc_2024)
print_accounts(2025, acc_2025)


def get_charge_accounts(acc):
    result = {}
    for k in sorted(acc.keys()):
        if get_class(k) == '6':
            net = acc[k]['debit'] - acc[k]['credit']
            if abs(net) > 0.001:
                result[k] = {'label': acc[k]['label'], 'amount': net}
    return result

def get_produit_accounts(acc):
    result = {}
    for k in sorted(acc.keys()):
        if get_class(k) == '7':
            net = acc[k]['credit'] - acc[k]['debit']
            if abs(net) > 0.001:
                result[k] = {'label': acc[k]['label'], 'amount': net}
    return result


charges_2023 = get_charge_accounts(acc_2023)
charges_2024 = get_charge_accounts(acc_2024)
charges_2025 = get_charge_accounts(acc_2025)

produits_2023 = get_produit_accounts(acc_2023)
produits_2024 = get_produit_accounts(acc_2024)
produits_2025 = get_produit_accounts(acc_2025)

total_charges = [
    sum(v['amount'] for v in charges_2023.values()),
    sum(v['amount'] for v in charges_2024.values()),
    sum(v['amount'] for v in charges_2025.values()),
]
total_produits = [
    sum(v['amount'] for v in produits_2023.values()),
    sum(v['amount'] for v in produits_2024.values()),
    sum(v['amount'] for v in produits_2025.values()),
]
resultats = [total_produits[i] - total_charges[i] for i in range(3)]

print(f"\n=== RÉSULTATS ===")
print(f"Charges  2023: {total_charges[0]:>12,.2f} €  |  2024: {total_charges[1]:>12,.2f} €  |  2025: {total_charges[2]:>12,.2f} €")
print(f"Produits 2023: {total_produits[0]:>12,.2f} €  |  2024: {total_produits[1]:>12,.2f} €  |  2025: {total_produits[2]:>12,.2f} €")
print(f"Résultat 2023: {resultats[0]:>12,.2f} €  |  2024: {resultats[1]:>12,.2f} €  |  2025: {resultats[2]:>12,.2f} €")

print(f"\n=== DÉTAIL CHARGES 2023 ===")
for k, v in sorted(charges_2023.items()):
    print(f"  {k:12s} {v['label'][:40]:40s}  {v['amount']:>12,.2f}")

print(f"\n=== DÉTAIL CHARGES 2024 ===")
for k, v in sorted(charges_2024.items()):
    print(f"  {k:12s} {v['label'][:40]:40s}  {v['amount']:>12,.2f}")

print(f"\n=== DÉTAIL PRODUITS 2023 ===")
for k, v in sorted(produits_2023.items()):
    print(f"  {k:12s} {v['label'][:40]:40s}  {v['amount']:>12,.2f}")

print(f"\n=== DÉTAIL PRODUITS 2024 ===")
for k, v in sorted(produits_2024.items()):
    print(f"  {k:12s} {v['label'][:40]:40s}  {v['amount']:>12,.2f}")


# =====================================================================
# 5. GÉNÉRATION EXCEL
# =====================================================================

wb = openpyxl.Workbook()

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
charges_fill = PatternFill(start_color='F4B084', end_color='F4B084', fill_type='solid')
produits_fill = PatternFill(start_color='A9D18E', end_color='A9D18E', fill_type='solid')

thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
euro_fmt = '#,##0.00 €'

def style_header(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin_border

def style_section(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = section_font
        cell.fill = section_fill
        cell.border = thin_border

def wr(ws, row, label, v23, v24, v25, indent=0, bold=False, fill=None):
    cell_a = ws.cell(row=row, column=1, value=('   ' * indent) + label)
    cell_a.font = total_font if bold else normal_font
    cell_a.border = thin_border
    cell_a.alignment = Alignment(wrap_text=True)
    if fill:
        cell_a.fill = fill
    for ci, val in enumerate([v23, v24, v25], 2):
        cell = ws.cell(row=row, column=ci, value=val)
        if val is not None:
            cell.number_format = euro_fmt
        cell.font = total_font if bold else normal_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='right')
        if fill:
            cell.fill = fill

def best_label(key, *dicts):
    for d in dicts:
        if key in d:
            return d[key].get('label', key)
    return key

subclass_labels = {
    '60': 'ACHATS ET FLUIDES',
    '61': 'SERVICES EXTÉRIEURS',
    '62': 'AUTRES SERVICES EXTÉRIEURS',
    '66': 'CHARGES FINANCIÈRES',
    '67': 'CHARGES EXCEPTIONNELLES / TRAVAUX AG',
    '69': 'CHARGES ET FRAIS PRIVATIFS',
    '70': 'PROVISIONS SUR CHARGES',
    '71': 'PRODUITS EXCEPTIONNELS',
}

all_charge_keys = sorted(set(list(charges_2023.keys()) + list(charges_2024.keys()) + list(charges_2025.keys())))
all_produit_keys = sorted(set(list(produits_2023.keys()) + list(produits_2024.keys()) + list(produits_2025.keys())))


# --- FEUILLE 1 : COMPTE DE RÉSULTAT ---
ws1 = wb.active
ws1.title = "Compte de Résultat"
ws1.column_dimensions['A'].width = 60
ws1.column_dimensions['B'].width = 18
ws1.column_dimensions['C'].width = 18
ws1.column_dimensions['D'].width = 18

r = 1
ws1.merge_cells('A1:D1')
ws1.cell(row=r, column=1, value="COMPTE DE RÉSULTAT COMPARATIF PAR NATURE DE COMPTE").font = title_font
ws1.cell(row=r, column=1).alignment = Alignment(horizontal='center')
r += 1
ws1.merge_cells('A2:D2')
ws1.cell(row=r, column=1, value="SDC 17 Rue Léon Cladel, Sèvres - Copropriété PERDUE").font = Font(italic=True, size=11)
ws1.cell(row=r, column=1).alignment = Alignment(horizontal='center')
r += 1
ws1.merge_cells('A3:D3')
ws1.cell(row=r, column=1, value="Basé sur les mouvements de l'exercice (hors à-nouveaux, clôtures et décomptes)").font = Font(italic=True, size=10, color='2F5496')
ws1.cell(row=r, column=1).alignment = Alignment(horizontal='center')
r += 2

# CHARGES
ws1.cell(row=r, column=1, value="CHARGES (Classe 6)")
ws1.cell(row=r, column=2, value="2023\n(01/01 - 31/12)")
ws1.cell(row=r, column=3, value="2024\n(01/01 - 31/12)")
ws1.cell(row=r, column=4, value="2025\n(01/01 - 20/07)*")
style_header(ws1, r, 4)
r += 1

cur_sc = None
sc_tot = [0, 0, 0]

for key in all_charge_keys:
    sc = get_subclass(key)
    if sc != cur_sc:
        if cur_sc is not None:
            wr(ws1, r, f"Sous-total {cur_sc} - {subclass_labels.get(cur_sc, '')}", sc_tot[0], sc_tot[1], sc_tot[2], bold=True, fill=subtotal_fill)
            r += 2
        cur_sc = sc
        sc_tot = [0, 0, 0]
        style_section(ws1, r, 4)
        ws1.cell(row=r, column=1, value=f"{sc} - {subclass_labels.get(sc, 'AUTRES')}")
        r += 1

    lbl = best_label(key, charges_2023, charges_2024, charges_2025)
    v23 = charges_2023.get(key, {}).get('amount', 0)
    v24 = charges_2024.get(key, {}).get('amount', 0)
    v25 = charges_2025.get(key, {}).get('amount', 0)
    wr(ws1, r, f"{key} - {lbl}", v23 or 0, v24 or 0, v25 or 0, indent=1)
    sc_tot[0] += v23; sc_tot[1] += v24; sc_tot[2] += v25
    r += 1

if cur_sc:
    wr(ws1, r, f"Sous-total {cur_sc} - {subclass_labels.get(cur_sc, '')}", sc_tot[0], sc_tot[1], sc_tot[2], bold=True, fill=subtotal_fill)
    r += 2

wr(ws1, r, "TOTAL CHARGES (Classe 6)", total_charges[0], total_charges[1], total_charges[2], bold=True, fill=charges_fill)
for c in range(1, 5):
    ws1.cell(row=r, column=c).font = Font(bold=True, size=11)
r += 3

# PRODUITS
ws1.cell(row=r, column=1, value="PRODUITS (Classe 7)")
ws1.cell(row=r, column=2, value="2023")
ws1.cell(row=r, column=3, value="2024")
ws1.cell(row=r, column=4, value="2025*")
style_header(ws1, r, 4)
r += 1

cur_sc = None
sc_tot = [0, 0, 0]

for key in all_produit_keys:
    sc = get_subclass(key)
    if sc != cur_sc:
        if cur_sc is not None:
            wr(ws1, r, f"Sous-total {cur_sc} - {subclass_labels.get(cur_sc, '')}", sc_tot[0], sc_tot[1], sc_tot[2], bold=True, fill=subtotal_fill)
            r += 2
        cur_sc = sc
        sc_tot = [0, 0, 0]
        style_section(ws1, r, 4)
        ws1.cell(row=r, column=1, value=f"{sc} - {subclass_labels.get(sc, 'AUTRES')}")
        r += 1

    lbl = best_label(key, produits_2023, produits_2024, produits_2025)
    v23 = produits_2023.get(key, {}).get('amount', 0)
    v24 = produits_2024.get(key, {}).get('amount', 0)
    v25 = produits_2025.get(key, {}).get('amount', 0)
    wr(ws1, r, f"{key} - {lbl}", v23 or 0, v24 or 0, v25 or 0, indent=1)
    sc_tot[0] += v23; sc_tot[1] += v24; sc_tot[2] += v25
    r += 1

if cur_sc:
    wr(ws1, r, f"Sous-total {cur_sc} - {subclass_labels.get(cur_sc, '')}", sc_tot[0], sc_tot[1], sc_tot[2], bold=True, fill=subtotal_fill)
    r += 2

wr(ws1, r, "TOTAL PRODUITS (Classe 7)", total_produits[0], total_produits[1], total_produits[2], bold=True, fill=produits_fill)
for c in range(1, 5):
    ws1.cell(row=r, column=c).font = Font(bold=True, size=11)
r += 3

# RÉSULTAT
ws1.cell(row=r, column=1, value="RÉSULTAT DE L'EXERCICE")
ws1.cell(row=r, column=2, value="2023"); ws1.cell(row=r, column=3, value="2024"); ws1.cell(row=r, column=4, value="2025*")
style_header(ws1, r, 4)
r += 1
wr(ws1, r, "Résultat = Produits - Charges", resultats[0], resultats[1], resultats[2], bold=True, fill=result_fill)
for c in range(1, 5):
    ws1.cell(row=r, column=c).font = Font(bold=True, size=12)
r += 2
ws1.cell(row=r, column=1, value="Résultat positif = excédent | Résultat négatif = déficit").font = Font(italic=True, size=9)
r += 1
ws1.cell(row=r, column=1, value="* 2025 : exercice en cours au 20/07/2025, données partielles.").font = Font(italic=True, size=9, color='FF0000')


# --- FEUILLE 2 : BILAN CLASSIQUE ACTIF / PASSIF ---
ws2 = wb.create_sheet("Bilan Comparatif")

# Colonnes : A=Actif libellé, B/C/D=montants 23/24/25, E=séparateur,
#            F=Passif libellé, G/H/I=montants 23/24/25
ws2.column_dimensions['A'].width = 42
ws2.column_dimensions['B'].width = 15
ws2.column_dimensions['C'].width = 15
ws2.column_dimensions['D'].width = 15
ws2.column_dimensions['E'].width = 3
ws2.column_dimensions['F'].width = 42
ws2.column_dimensions['G'].width = 15
ws2.column_dimensions['H'].width = 15
ws2.column_dimensions['I'].width = 15

# Couleurs spécifiques bilan
actif_header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
passif_header_fill = PatternFill(start_color='833C0B', end_color='833C0B', fill_type='solid')
actif_section_fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
passif_section_fill = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
actif_total_fill = PatternFill(start_color='9DC3E6', end_color='9DC3E6', fill_type='solid')
passif_total_fill = PatternFill(start_color='F4B084', end_color='F4B084', fill_type='solid')
grand_total_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
resultat_pos_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
resultat_neg_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
sep_fill = PatternFill(start_color='404040', end_color='404040', fill_type='solid')

white_font_b = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
white_font_big = Font(name='Calibri', bold=True, size=12, color='FFFFFF')

def bilan_cell(ws, row, col, value, font=normal_font, fill=None, fmt=None, align='left'):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font
    cell.border = thin_border
    cell.alignment = Alignment(horizontal=align, wrap_text=True)
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    return cell

def bilan_sep(ws, row):
    cell = ws.cell(row=row, column=5)
    cell.fill = sep_fill
    cell.border = thin_border

# =====================================================================
# BILAN : Données des soldes de clôture vérifiées
# (issues des grands livres, lignes "Cloture Exercice" et à-nouveaux)
# =====================================================================
# Pour le bilan, on utilise les soldes de fin d'exercice (vérifiés),
# car le parsing des mouvements ne permet pas de reconstituer les soldes
# cumulés de manière fiable.

actif_data = [
    ("Actif immobilisé", [
        ("Solde en attente travaux (1200)", [120.00, 0, 0]),
    ]),
    ("Créances copropriétaires (450)", [
        ("Copropriétaires - soldes débiteurs", [0, 0, 3776.31]),
    ]),
    ("Créances fournisseurs & tiers", [
        ("Fournisseurs débiteurs (4010)", [0, 0, 278.48]),
    ]),
    ("Comptes de régularisation - Actif", [
        ("Régularisation charges débiteur (4711)", [18665.97, 26363.21, 26363.21]),
        ("Régularisation travaux débiteur (4712)", [0, 911.97, 911.97]),
        ("Rompus débiteurs (4730)", [0.06, 0.09, 0.09]),
    ]),
    ("Trésorerie", [
        ("Compte courant Montepaschi", [2968.03, 9525.90, 13074.46]),
        ("Livret A Monte Paschi", [910.22, 1742.55, 1967.53]),
    ]),
]

passif_data = [
    ("Avances et fonds de travaux", [
        ("Avances de trésorerie (1031)", [2000.00, 2000.00, 2000.00]),
        ("Fonds de travaux (1050)", [910.22, 1742.55, 2256.18]),
    ]),
    ("Dettes fournisseurs", [
        ("Fournisseurs créditeurs (4010)", [1330.06, 659.45, 0]),
        ("Factures non parvenues (4080)", [0, 141.72, 0]),
    ]),
    ("Dettes et tiers créditeurs", [
        ("Anciens copropriétaires (4630)", [424.00, 0, 0]),
    ]),
    ("Comptes de régularisation - Passif", [
        ("Régularisation charges créditeur (4721)", [18000.00, 34000.00, 34000.00]),
    ]),
]

def sum_section(section):
    t = [0, 0, 0]
    for _, vals in section:
        for i in range(3):
            t[i] += vals[i]
    return t

actif_totals = [0, 0, 0]
for sec_name, entries in actif_data:
    st = sum_section(entries)
    for i in range(3):
        actif_totals[i] += st[i]

passif_totals = [0, 0, 0]
for sec_name, entries in passif_data:
    st = sum_section(entries)
    for i in range(3):
        passif_totals[i] += st[i]

# Pour 2023/2024 (clôturés), le résultat est déjà dans les soldes 4711/4721.
# Pour 2025 (en cours), on ajoute le résultat des classes 6/7 non encore soldées.
resultat_bilan = [0, 0, resultats[2]]
passif_plus_res = [passif_totals[i] + resultat_bilan[i] for i in range(3)]

# Build rows for display
def build_bilan_rows(data):
    rows = []
    for sec_name, entries in data:
        st = sum_section(entries)
        rows.append(('section', sec_name, [0, 0, 0]))
        for lbl, vals in entries:
            rows.append(('line', lbl, vals))
        rows.append(('subtotal', f"Total {sec_name}", st))
    return rows

actif_rows = build_bilan_rows(actif_data)
passif_rows = build_bilan_rows(passif_data)

# Titre
r = 1
ws2.merge_cells('A1:I1')
bilan_cell(ws2, r, 1, "BILAN COMPARATIF", font=title_font, align='center')
r += 1
ws2.merge_cells('A2:I2')
bilan_cell(ws2, r, 1, "SDC 17 Rue Léon Cladel, Sèvres - Copropriété PERDUE", font=Font(italic=True, size=11), align='center')
r += 2

# En-têtes ACTIF / PASSIF
for col in [1, 2, 3, 4]:
    bilan_cell(ws2, r, col, ["ACTIF", "2023", "2024", "2025*"][col-1], font=white_font_b, fill=actif_header_fill, align='center')
bilan_sep(ws2, r)
for col in [6, 7, 8, 9]:
    bilan_cell(ws2, r, col, ["PASSIF", "2023", "2024", "2025*"][col-6], font=white_font_b, fill=passif_header_fill, align='center')
r += 1

# Nombre max de lignes
max_rows = max(len(actif_rows), len(passif_rows))

def write_bilan_line(ws, row, col_start, row_data, side='actif'):
    if row_data is None:
        for c in range(col_start, col_start + 4):
            bilan_cell(ws, row, c, '')
        return

    kind, label, vals = row_data
    sec_fill = actif_section_fill if side == 'actif' else passif_section_fill
    tot_fill = actif_total_fill if side == 'actif' else passif_total_fill

    if kind == 'section':
        for c in range(col_start, col_start + 4):
            bilan_cell(ws, row, c, label if c == col_start else '', font=section_font, fill=sec_fill)
    elif kind == 'subtotal':
        bilan_cell(ws, row, col_start, label, font=total_font, fill=tot_fill)
        for i in range(3):
            bilan_cell(ws, row, col_start + 1 + i, vals[i], font=total_font, fill=tot_fill, fmt=euro_fmt, align='right')
    else:
        bilan_cell(ws, row, col_start, '   ' + label, font=normal_font)
        for i in range(3):
            bilan_cell(ws, row, col_start + 1 + i, vals[i] if abs(vals[i]) > 0.001 else 0, font=normal_font, fmt=euro_fmt, align='right')


start_r = r
for idx in range(max_rows):
    a_row = actif_rows[idx] if idx < len(actif_rows) else None
    p_row = passif_rows[idx] if idx < len(passif_rows) else None
    write_bilan_line(ws2, r, 1, a_row, 'actif')
    bilan_sep(ws2, r)
    write_bilan_line(ws2, r, 6, p_row, 'passif')
    r += 1

# Ajouter le résultat net au passif pour équilibrer
r += 1
bilan_cell(ws2, r, 1, '', font=normal_font)
bilan_cell(ws2, r, 2, '', font=normal_font)
bilan_cell(ws2, r, 3, '', font=normal_font)
bilan_cell(ws2, r, 4, '', font=normal_font)
bilan_sep(ws2, r)

res_label = "Résultat de l'exercice"
for i, res_val in enumerate(resultats):
    res_fill = resultat_pos_fill if res_val >= 0 else resultat_neg_fill
    sign = "excédent" if res_val >= 0 else "déficit"

bilan_cell(ws2, r, 6, f"Résultat net de l'exercice", font=Font(bold=True, size=10, color='006100'), fill=resultat_pos_fill)
for i in range(3):
    res_fill = resultat_pos_fill if resultats[i] >= 0 else resultat_neg_fill
    bilan_cell(ws2, r, 7 + i, resultat_bilan[i], font=Font(bold=True, size=10), fill=res_fill, fmt=euro_fmt, align='right')
r += 1

# Ligne vide séparatrice
bilan_sep(ws2, r)
r += 1

# TOTAUX GÉNÉRAUX ÉQUILIBRÉS

bilan_cell(ws2, r, 1, "TOTAL ACTIF", font=white_font_big, fill=grand_total_fill, align='center')
for i in range(3):
    bilan_cell(ws2, r, 2 + i, actif_totals[i], font=white_font_big, fill=grand_total_fill, fmt=euro_fmt, align='right')
bilan_sep(ws2, r)
bilan_cell(ws2, r, 6, "TOTAL PASSIF", font=white_font_big, fill=grand_total_fill, align='center')
for i in range(3):
    bilan_cell(ws2, r, 7 + i, passif_plus_res[i], font=white_font_big, fill=grand_total_fill, fmt=euro_fmt, align='right')
r += 2

# Vérification équilibre
for i, yr in enumerate(['2023', '2024', '2025']):
    diff = abs(actif_totals[i] - passif_plus_res[i])
    status = "✓ Équilibré" if diff < 0.02 else f"✗ Écart: {diff:.2f}"
    bilan_cell(ws2, r, 1, f"Contrôle {yr}: {status}", font=Font(italic=True, size=9, color='006100' if diff < 0.02 else 'FF0000'))
    r += 1

r += 1
ws2.cell(row=r, column=1, value="* 2025 : exercice en cours au 20/07/2025 (non clôturé).").font = Font(italic=True, size=9, color='FF0000')
r += 1
ws2.cell(row=r, column=1, value="Le résultat net est inscrit au passif pour équilibrer le bilan (excédent = passif / déficit = diminue le passif).").font = Font(italic=True, size=9)

print(f"\nBilan: Actif = {actif_totals}  Passif = {passif_totals}  Résultat = {resultats}")
print(f"Passif + Résultat = {passif_plus_res}")


# --- FEUILLE 3 : DÉTAIL ---
ws3 = wb.create_sheet("Détail Mouvements")
ws3.column_dimensions['A'].width = 14
ws3.column_dimensions['B'].width = 45
for col_letter in ['C','D','E','F','G','H','I','J','K']:
    ws3.column_dimensions[col_letter].width = 15

r = 1
ws3.merge_cells('A1:K1')
ws3.cell(row=r, column=1, value="DÉTAIL DES MOUVEMENTS PAR COMPTE").font = title_font
ws3.cell(row=r, column=1).alignment = Alignment(horizontal='center')
r += 2

headers_d = ["Compte", "Libellé", "Débit 23", "Crédit 23", "Net 23", "Débit 24", "Crédit 24", "Net 24", "Débit 25", "Crédit 25", "Net 25"]
for ci, h in enumerate(headers_d, 1):
    cell = ws3.cell(row=r, column=ci, value=h)
    cell.font = header_font; cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center', wrap_text=True); cell.border = thin_border
r += 1

all_keys = sorted(set(list(acc_2023.keys()) + list(acc_2024.keys()) + list(acc_2025.keys())))
for key in all_keys:
    lbl = best_label(key, acc_2023, acc_2024, acc_2025)
    ws3.cell(row=r, column=1, value=key).border = thin_border
    ws3.cell(row=r, column=2, value=lbl).border = thin_border
    ws3.cell(row=r, column=2).font = normal_font
    for yi, acc in enumerate([acc_2023, acc_2024, acc_2025]):
        cb = 3 + yi * 3
        d = acc.get(key, {}).get('debit', 0)
        c_v = acc.get(key, {}).get('credit', 0)
        s = d - c_v
        for ci2, v in enumerate([d, c_v, s]):
            cell = ws3.cell(row=r, column=cb + ci2, value=v if abs(v) > 0.001 else 0)
            cell.number_format = euro_fmt; cell.border = thin_border
            cell.alignment = Alignment(horizontal='right'); cell.font = normal_font
    r += 1


# --- SAVE ---
output_path = "/workspace/bilan_comptable_comparatif.xlsx"
wb.save(output_path)
print(f"\nFichier généré : {output_path}")
