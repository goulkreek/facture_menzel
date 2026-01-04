#!/usr/bin/env python3
"""
Script d'extraction du chiffre d'affaires depuis les factures PDF.
"""

import os
import re
import csv
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    print("Erreur: PyMuPDF non installé. Exécutez: pip3 install pymupdf")
    exit(1)

DOSSIER_FACTURES = Path("/Users/sroyer/Documents/Dev/maconnerieconcept/Factures_Menzel/FACTURE_2025")
CSV_OUTPUT = Path("/Users/sroyer/Documents/Dev/maconnerieconcept/Factures_Menzel/chiffre_affaires.csv")


def nettoyer_montant(montant_str: str) -> float:
    """Nettoie et convertit une chaîne de montant en float."""
    # Supprimer tous les types d'espaces (normaux, insécables, fins)
    cleaned = montant_str.replace(' ', '').replace('\u00a0', '').replace('\u202f', '')
    # Remplacer la virgule par un point
    cleaned = cleaned.replace(',', '.')
    return float(cleaned)


def extraire_montant(texte: str) -> dict:
    """
    Extrait les montants d'une facture.
    Retourne un dict avec numero_facture, date, client, montant_ht, tva, montant_ttc, type_facture
    """
    result = {
        'numero_facture': None,
        'date': None,
        'client': None,
        'montant_ht': None,
        'tva_taux': None,
        'tva_montant': None,
        'montant_ttc': None,
        'type_facture': None
    }

    # Numéro de facture
    match = re.search(r'FACTURE\s*N[°o]?\s*:?\s*(\d+[_/]?\d*)', texte, re.IGNORECASE)
    if match:
        result['numero_facture'] = match.group(1).replace('_', '-')

    # Date
    match = re.search(r'Date\s*:?\s*(\d{2}/\d{2}/\d{4})', texte)
    if match:
        result['date'] = match.group(1)

    # Client
    match = re.search(r'Client\s*:?\s*([^\n]+)', texte, re.IGNORECASE)
    if match:
        result['client'] = match.group(1).strip()

    # Détecter le type de facture
    if 'AUTOLIQUIDATION' in texte.upper():
        result['type_facture'] = 'Autoliquidation'
        # Montant HT pour autoliquidation (pas de TTC)
        # Pattern: "HT" suivi d'un montant (avec espaces insécables possibles)
        matches = re.findall(r'(?:^|\n)\s*HT\s*\n?\s*([\d\s\u00a0\u202f]+[,.]?\d*)\s*€', texte, re.IGNORECASE)
        if matches:
            result['montant_ht'] = nettoyer_montant(matches[-1])
            result['montant_ttc'] = result['montant_ht']  # Pas de TVA
    else:
        result['type_facture'] = 'Standard'

        # Total HT
        match = re.search(r'Total\s*HT\s*\n?\s*([\d\s\u00a0\u202f]+[,.]?\d*)\s*€', texte, re.IGNORECASE)
        if match:
            result['montant_ht'] = nettoyer_montant(match.group(1))

        # TVA
        match = re.search(r'TVA\s*\n?\s*([\d,]+)\s*%\s*\n?\s*([\d\s\u00a0\u202f]+[,.]?\d*)\s*€', texte, re.IGNORECASE)
        if match:
            result['tva_taux'] = match.group(1).replace(',', '.')
            result['tva_montant'] = nettoyer_montant(match.group(2))

        # Total TTC
        match = re.search(r'Total\s*TTC\s*\n?\s*([\d\s\u00a0\u202f]+[,.]?\d*)\s*€', texte, re.IGNORECASE)
        if match:
            result['montant_ttc'] = nettoyer_montant(match.group(1))

    return result


def traiter_pdf(chemin_pdf: Path) -> dict:
    """Traite un fichier PDF et extrait les informations de facturation."""
    try:
        doc = fitz.open(chemin_pdf)
        texte = ""
        for page in doc:
            texte += page.get_text()
        doc.close()

        result = extraire_montant(texte)
        result['fichier'] = chemin_pdf.name
        return result

    except Exception as e:
        print(f"Erreur lors du traitement de {chemin_pdf.name}: {e}")
        return None


def main():
    """Fonction principale."""
    print("=== Extraction du Chiffre d'Affaires ===\n")

    # Lister les PDFs de factures
    pdfs = list(DOSSIER_FACTURES.glob("*PDF*.pdf"))
    print(f"Nombre de factures PDF: {len(pdfs)}\n")

    resultats = []
    erreurs = []

    for i, pdf in enumerate(pdfs, 1):
        if i % 50 == 0:
            print(f"Traitement: {i}/{len(pdfs)}")

        result = traiter_pdf(pdf)
        if result and result['montant_ht'] is not None:
            resultats.append(result)
        else:
            erreurs.append(pdf.name)

    # Trier par numéro de facture
    def sort_key(r):
        num = r['numero_facture'] or '0'
        # Extraire le premier nombre
        match = re.match(r'(\d+)', num)
        return int(match.group(1)) if match else 0

    resultats.sort(key=sort_key)

    # Écrire le CSV
    fieldnames = ['numero_facture', 'date', 'client', 'type_facture', 'montant_ht', 'tva_taux', 'tva_montant', 'montant_ttc', 'fichier']

    with open(CSV_OUTPUT, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
        writer.writeheader()

        for r in resultats:
            row = {k: r.get(k, '') for k in fieldnames}
            # Format français pour les montants
            for key in ['montant_ht', 'tva_montant', 'montant_ttc']:
                if row[key]:
                    row[key] = f"{row[key]:.2f}".replace('.', ',')
            writer.writerow(row)

    print(f"\nCSV généré: {CSV_OUTPUT}")

    # Statistiques
    print(f"\n=== RÉSUMÉ ===")
    print(f"Factures traitées: {len(resultats)}")
    print(f"Erreurs: {len(erreurs)}")

    # Calcul du CA
    total_ht = sum(r['montant_ht'] for r in resultats if r['montant_ht'])
    total_ttc = sum(r['montant_ttc'] for r in resultats if r['montant_ttc'])
    total_tva = sum(r['tva_montant'] for r in resultats if r['tva_montant'])

    # Par type de facture
    autoliq = [r for r in resultats if r['type_facture'] == 'Autoliquidation']
    standard = [r for r in resultats if r['type_facture'] == 'Standard']

    autoliq_ht = sum(r['montant_ht'] for r in autoliq if r['montant_ht'])
    standard_ht = sum(r['montant_ht'] for r in standard if r['montant_ht'])
    standard_ttc = sum(r['montant_ttc'] for r in standard if r['montant_ttc'])

    print(f"\n=== CHIFFRE D'AFFAIRES 2025 ===")
    print(f"")
    print(f"TOTAL HT:              {total_ht:,.2f} €".replace(',', ' ').replace('.', ','))
    print(f"Total TVA:             {total_tva:,.2f} €".replace(',', ' ').replace('.', ','))
    print(f"TOTAL TTC:             {total_ttc:,.2f} €".replace(',', ' ').replace('.', ','))
    print(f"")
    print(f"--- Par type ---")
    print(f"Autoliquidation (sous-traitance): {len(autoliq)} factures = {autoliq_ht:,.2f} € HT".replace(',', ' ').replace('.', ','))
    print(f"Standard (particuliers):          {len(standard)} factures = {standard_ht:,.2f} € HT / {standard_ttc:,.2f} € TTC".replace(',', ' ').replace('.', ','))

    # Facture min/max
    if resultats:
        min_facture = min(resultats, key=lambda r: r['montant_ht'] or 0)
        max_facture = max(resultats, key=lambda r: r['montant_ht'] or 0)
        moyenne = total_ht / len(resultats)

        print(f"\n--- Statistiques ---")
        print(f"Facture min: {min_facture['montant_ht']:.2f} € (N°{min_facture['numero_facture']})".replace('.', ','))
        print(f"Facture max: {max_facture['montant_ht']:.2f} € (N°{max_facture['numero_facture']})".replace('.', ','))
        print(f"Moyenne:     {moyenne:.2f} €".replace('.', ','))

    if erreurs:
        print(f"\n--- Fichiers en erreur ---")
        for e in erreurs[:10]:
            print(f"  - {e}")
        if len(erreurs) > 10:
            print(f"  ... et {len(erreurs) - 10} autres")


if __name__ == "__main__":
    main()
