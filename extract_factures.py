#!/usr/bin/env python3
"""
Script d'extraction des données des factures PDF
Entreprise: SASU Maçonnerie Concept
"""

import os
import re
import fitz  # PyMuPDF
import csv
from datetime import datetime
from collections import defaultdict

FACTURES_DIR = "/Users/sroyer/Documents/Dev/maconnerieconcept/Factures_Menzel/FACTURE_2025"
OUTPUT_CSV = "/Users/sroyer/Documents/Dev/maconnerieconcept/Factures_Menzel/analyse_factures.csv"

def extract_invoice_data(pdf_path):
    """Extrait les données d'une facture PDF"""
    try:
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()

        # Extraire le numéro de facture
        num_match = re.search(r'FACTURE\s*N[°º]?\s*:?\s*(\d+)[_\-]?(\d{4})?', text, re.IGNORECASE)
        num_facture = num_match.group(1) if num_match else None

        # Extraire la date
        date_match = re.search(r'Date\s*:?\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})', text, re.IGNORECASE)
        date_str = date_match.group(1) if date_match else None

        # Parser la date
        date_obj = None
        if date_str:
            try:
                # Essayer différents formats
                for fmt in ['%d/%m/%Y', '%d-%m-%Y', '%d/%m/%y']:
                    try:
                        date_obj = datetime.strptime(date_str, fmt)
                        break
                    except ValueError:
                        continue
            except:
                pass

        # Extraire le client principal (Leroy Merlin ou autre)
        client_principal = "Particulier"
        if "LEROY MERLIN" in text.upper() or "LM " in text.upper():
            # Identifier le magasin LM
            if "LA VALETTE" in text.upper():
                client_principal = "Leroy Merlin La Valette"
            elif "AUBAGNE" in text.upper():
                client_principal = "Leroy Merlin Aubagne"
            elif "CABRIES" in text.upper() or "CAB" in text.upper():
                client_principal = "Leroy Merlin Cabriès"
            elif "MARTIGUES" in text.upper() or "MART" in text.upper():
                client_principal = "Leroy Merlin Martigues"
            elif "TOULON" in text.upper():
                client_principal = "Leroy Merlin Toulon"
            elif "GRAND LITTORAL" in text.upper() or "GR LIT" in text.upper():
                client_principal = "Leroy Merlin Grand Littoral"
            else:
                client_principal = "Leroy Merlin"
        elif "LAPEY" in text.upper():
            client_principal = "Lapeyre"
        elif "CASTO" in text.upper():
            client_principal = "Castorama"
        elif "DISTRILAP" in text.upper():
            client_principal = "Distrilap"

        # Extraire le client final (REF CLIENT)
        ref_client_match = re.search(r'REF\s*CLIENT\s*:?\s*([^\n]+)', text, re.IGNORECASE)
        ref_client = ref_client_match.group(1).strip() if ref_client_match else ""

        # Extraire le montant HT (chercher le total final)
        montant_ht = None

        # Chercher "HT" suivi d'un montant
        ht_patterns = [
            r'HT\s+(\d[\d\s]*[,\.]\d{2})\s*€',
            r'HT\s*:?\s*(\d[\d\s]*[,\.]\d{2})\s*€',
            r'Total\s*HT\s*:?\s*(\d[\d\s]*[,\.]\d{2})\s*€',
            r'(\d[\d\s]*[,\.]\d{2})\s*€\s*$',
        ]

        for pattern in ht_patterns:
            ht_match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
            if ht_match:
                montant_str = ht_match.group(1).replace(' ', '').replace(',', '.')
                try:
                    montant_ht = float(montant_str)
                    break
                except ValueError:
                    continue

        # Si toujours pas trouvé, chercher le dernier montant en euros
        if montant_ht is None:
            all_amounts = re.findall(r'(\d[\d\s]*[,\.]\d{2})\s*€', text)
            if all_amounts:
                montant_str = all_amounts[-1].replace(' ', '').replace(',', '.')
                try:
                    montant_ht = float(montant_str)
                except ValueError:
                    pass

        # Extraire la nature des travaux
        nature_match = re.search(r'Nature\s*des\s*travaux\s*:?\s*([^\n]+(?:\n[^\n]+)?)', text, re.IGNORECASE)
        nature_travaux = nature_match.group(1).strip().replace('\n', ' ') if nature_match else ""

        return {
            'numero': num_facture,
            'date': date_obj,
            'date_str': date_str,
            'client_principal': client_principal,
            'ref_client': ref_client,
            'montant_ht': montant_ht,
            'nature_travaux': nature_travaux,
            'fichier': os.path.basename(pdf_path)
        }

    except Exception as e:
        print(f"Erreur avec {pdf_path}: {e}")
        return None

def main():
    """Extrait les données de toutes les factures"""
    factures = []

    # Lister tous les fichiers PDF
    pdf_files = [f for f in os.listdir(FACTURES_DIR) if f.endswith('.pdf')]
    pdf_files.sort()

    print(f"Extraction de {len(pdf_files)} factures...")

    for i, pdf_file in enumerate(pdf_files):
        pdf_path = os.path.join(FACTURES_DIR, pdf_file)
        data = extract_invoice_data(pdf_path)
        if data:
            factures.append(data)

        if (i + 1) % 50 == 0:
            print(f"  {i + 1}/{len(pdf_files)} factures traitées...")

    print(f"\n{len(factures)} factures extraites avec succès")

    # Sauvegarder en CSV
    with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=[
            'numero', 'date_str', 'client_principal', 'ref_client',
            'montant_ht', 'nature_travaux', 'fichier'
        ], delimiter=';')
        writer.writeheader()
        for facture in factures:
            row = {k: v for k, v in facture.items() if k != 'date'}
            writer.writerow(row)

    print(f"Données exportées vers: {OUTPUT_CSV}")

    # Afficher les statistiques
    print("\n" + "="*60)
    print("ANALYSE DE L'ACTIVITÉ - MAÇONNERIE CONCEPT")
    print("="*60)

    # Filtrer les factures avec montant valide
    factures_valides = [f for f in factures if f['montant_ht'] and f['montant_ht'] > 0]

    # Chiffre d'affaires total
    ca_total = sum(f['montant_ht'] for f in factures_valides)
    print(f"\nCHIFFRE D'AFFAIRES TOTAL: {ca_total:,.2f} € HT".replace(',', ' '))
    print(f"Nombre de factures: {len(factures_valides)}")
    print(f"Montant moyen par facture: {ca_total/len(factures_valides):,.2f} €".replace(',', ' '))

    # CA par mois
    ca_par_mois = defaultdict(float)
    nb_par_mois = defaultdict(int)
    for f in factures_valides:
        if f['date']:
            key = f['date'].strftime('%Y-%m')
            ca_par_mois[key] += f['montant_ht']
            nb_par_mois[key] += 1

    print("\n" + "-"*60)
    print("CHIFFRE D'AFFAIRES PAR MOIS:")
    print("-"*60)
    for mois in sorted(ca_par_mois.keys()):
        ca = ca_par_mois[mois]
        nb = nb_par_mois[mois]
        print(f"  {mois}: {ca:>10,.2f} € ({nb} factures)".replace(',', ' '))

    # CA par client principal
    ca_par_client = defaultdict(float)
    nb_par_client = defaultdict(int)
    for f in factures_valides:
        ca_par_client[f['client_principal']] += f['montant_ht']
        nb_par_client[f['client_principal']] += 1

    print("\n" + "-"*60)
    print("CHIFFRE D'AFFAIRES PAR CLIENT PRINCIPAL:")
    print("-"*60)
    for client, ca in sorted(ca_par_client.items(), key=lambda x: -x[1]):
        nb = nb_par_client[client]
        pct = (ca / ca_total) * 100
        print(f"  {client:<30}: {ca:>10,.2f} € ({nb:>3} factures, {pct:>5.1f}%)".replace(',', ' '))

    # Statistiques montants
    montants = [f['montant_ht'] for f in factures_valides]
    print("\n" + "-"*60)
    print("STATISTIQUES DES MONTANTS:")
    print("-"*60)
    print(f"  Montant minimum: {min(montants):,.2f} €".replace(',', ' '))
    print(f"  Montant maximum: {max(montants):,.2f} €".replace(',', ' '))
    print(f"  Montant moyen: {sum(montants)/len(montants):,.2f} €".replace(',', ' '))
    print(f"  Montant médian: {sorted(montants)[len(montants)//2]:,.2f} €".replace(',', ' '))

if __name__ == "__main__":
    main()
