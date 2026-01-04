#!/usr/bin/env python3
"""
Génère un rapport PDF de l'activité Maçonnerie Concept
"""

import fitz  # PyMuPDF
from datetime import datetime

OUTPUT_PDF = "/Users/sroyer/Documents/Dev/maconnerieconcept/Factures_Menzel/RAPPORT_ACTIVITE_2025.pdf"

def create_pdf_report():
    """Crée le rapport PDF"""

    # Créer un nouveau document
    doc = fitz.open()

    # Couleurs
    BLEU = (0.2, 0.4, 0.7)
    GRIS = (0.3, 0.3, 0.3)
    NOIR = (0, 0, 0)

    # ===== PAGE 1 : Couverture et Synthèse =====
    page = doc.new_page(width=595, height=842)  # A4

    # Titre principal
    page.insert_text((50, 80), "RAPPORT D'ANALYSE", fontsize=28, fontname="helv", color=BLEU)
    page.insert_text((50, 115), "D'ACTIVITÉ 2025", fontsize=28, fontname="helv", color=BLEU)

    # Ligne de séparation
    page.draw_line((50, 130), (545, 130), color=BLEU, width=2)

    # Nom entreprise
    page.insert_text((50, 170), "SASU MAÇONNERIE CONCEPT", fontsize=20, fontname="helv", color=NOIR)
    page.insert_text((50, 195), "Spécialiste de la fenêtre de toit", fontsize=14, fontname="helv", color=GRIS)

    # Informations entreprise
    y = 240
    infos = [
        ("Dirigeant", "Mr Menzel Franck"),
        ("Siège social", "208 Chemin de la croix des signaux"),
        ("", "83740 La Cadière d'Azur"),
        ("Téléphone", "06 12 36 03 48"),
        ("SIRET", "494 384 456 00039"),
        ("Certification RGE", "N° 105446"),
    ]

    for label, value in infos:
        if label:
            page.insert_text((50, y), f"{label}:", fontsize=10, fontname="helv", color=GRIS)
            page.insert_text((180, y), value, fontsize=10, fontname="helv", color=NOIR)
        else:
            page.insert_text((180, y), value, fontsize=10, fontname="helv", color=NOIR)
        y += 18

    # Encadré synthèse
    y = 380
    page.draw_rect(fitz.Rect(50, y, 545, y + 150), color=BLEU, width=2)
    page.insert_text((60, y + 25), "SYNTHÈSE FINANCIÈRE 2025", fontsize=14, fontname="helv", color=BLEU)

    page.insert_text((60, y + 55), "Chiffre d'Affaires (extrait):", fontsize=11, fontname="helv", color=GRIS)
    page.insert_text((280, y + 55), "129 307 € HT", fontsize=14, fontname="helv", color=NOIR)

    page.insert_text((60, y + 80), "CA estimé (total):", fontsize=11, fontname="helv", color=GRIS)
    page.insert_text((280, y + 80), "~162 000 € HT", fontsize=14, fontname="helv", color=NOIR)

    page.insert_text((60, y + 105), "Nombre de factures:", fontsize=11, fontname="helv", color=GRIS)
    page.insert_text((280, y + 105), "310", fontsize=14, fontname="helv", color=NOIR)

    page.insert_text((60, y + 130), "Montant moyen / facture:", fontsize=11, fontname="helv", color=GRIS)
    page.insert_text((280, y + 130), "524 €", fontsize=14, fontname="helv", color=NOIR)

    # Encadré répartition
    y = 560
    page.draw_rect(fitz.Rect(50, y, 545, y + 100), color=BLEU, width=1)
    page.insert_text((60, y + 25), "RÉPARTITION DU CHIFFRE D'AFFAIRES", fontsize=12, fontname="helv", color=BLEU)

    page.insert_text((60, y + 55), "Sous-traitance (Leroy Merlin, Castorama...):", fontsize=10, fontname="helv", color=GRIS)
    page.insert_text((320, y + 55), "74 552 € (57,7%)", fontsize=10, fontname="helv", color=NOIR)

    page.insert_text((60, y + 75), "Particuliers:", fontsize=10, fontname="helv", color=GRIS)
    page.insert_text((320, y + 75), "54 755 € (42,3%)", fontsize=10, fontname="helv", color=NOIR)

    # Footer
    page.insert_text((50, 800), f"Rapport généré le {datetime.now().strftime('%d/%m/%Y')}",
                     fontsize=8, fontname="helv", color=GRIS)
    page.insert_text((450, 800), "Page 1/3", fontsize=8, fontname="helv", color=GRIS)

    # ===== PAGE 2 : Analyse détaillée =====
    page = doc.new_page(width=595, height=842)

    # Titre
    page.insert_text((50, 50), "ANALYSE DÉTAILLÉE", fontsize=18, fontname="helv", color=BLEU)
    page.draw_line((50, 60), (545, 60), color=BLEU, width=1)

    # CA par trimestre
    y = 90
    page.insert_text((50, y), "Chiffre d'Affaires par Trimestre", fontsize=12, fontname="helv", color=BLEU)
    y += 25

    trimestres = [
        ("T1 2025 (Jan-Mar)", "28 779 €", "55 factures", "-"),
        ("T2 2025 (Avr-Jun)", "29 196 €", "57 factures", "+1,4%"),
        ("T3 2025 (Jul-Sep)", "31 908 €", "64 factures", "+9,3%"),
        ("T4 2025 (Oct-Déc)", "39 424 €", "71 factures", "+23,5%"),
    ]

    # En-tête tableau
    page.insert_text((50, y), "Période", fontsize=9, fontname="helv", color=GRIS)
    page.insert_text((200, y), "CA HT", fontsize=9, fontname="helv", color=GRIS)
    page.insert_text((300, y), "Nb Factures", fontsize=9, fontname="helv", color=GRIS)
    page.insert_text((420, y), "Évolution", fontsize=9, fontname="helv", color=GRIS)
    y += 5
    page.draw_line((50, y), (500, y), color=GRIS, width=0.5)
    y += 15

    for periode, ca, nb, evol in trimestres:
        page.insert_text((50, y), periode, fontsize=9, fontname="helv", color=NOIR)
        page.insert_text((200, y), ca, fontsize=9, fontname="helv", color=NOIR)
        page.insert_text((300, y), nb, fontsize=9, fontname="helv", color=NOIR)
        color = (0, 0.5, 0) if "+" in evol else NOIR
        page.insert_text((420, y), evol, fontsize=9, fontname="helv", color=color)
        y += 18

    # CA par mois
    y += 20
    page.insert_text((50, y), "Chiffre d'Affaires par Mois", fontsize=12, fontname="helv", color=BLEU)
    y += 25

    mois_data = [
        ("Janvier", "8 635 €", "15"), ("Février", "11 535 €", "22"), ("Mars", "8 609 €", "18"),
        ("Avril", "10 923 €", "23"), ("Mai", "9 433 €", "18"), ("Juin", "8 840 €", "16"),
        ("Juillet", "12 359 €", "25"), ("Août", "6 604 €", "15"), ("Septembre", "12 945 €", "24"),
        ("Octobre", "14 927 €", "26"), ("Novembre", "16 525 €", "28"), ("Décembre", "7 972 €", "17"),
    ]

    page.insert_text((50, y), "Mois", fontsize=9, fontname="helv", color=GRIS)
    page.insert_text((150, y), "CA HT", fontsize=9, fontname="helv", color=GRIS)
    page.insert_text((250, y), "Factures", fontsize=9, fontname="helv", color=GRIS)
    page.insert_text((320, y), "Mois", fontsize=9, fontname="helv", color=GRIS)
    page.insert_text((420, y), "CA HT", fontsize=9, fontname="helv", color=GRIS)
    page.insert_text((500, y), "Fact.", fontsize=9, fontname="helv", color=GRIS)
    y += 5
    page.draw_line((50, y), (545, y), color=GRIS, width=0.5)
    y += 15

    for i in range(6):
        m1, ca1, nb1 = mois_data[i]
        m2, ca2, nb2 = mois_data[i + 6]
        page.insert_text((50, y), m1, fontsize=9, fontname="helv", color=NOIR)
        page.insert_text((150, y), ca1, fontsize=9, fontname="helv", color=NOIR)
        page.insert_text((250, y), nb1, fontsize=9, fontname="helv", color=NOIR)
        page.insert_text((320, y), m2, fontsize=9, fontname="helv", color=NOIR)
        page.insert_text((420, y), ca2, fontsize=9, fontname="helv", color=NOIR)
        page.insert_text((500, y), nb2, fontsize=9, fontname="helv", color=NOIR)
        y += 18

    # Clients principaux
    y += 30
    page.insert_text((50, y), "Répartition par Client", fontsize=12, fontname="helv", color=BLEU)
    y += 25

    clients = [
        ("Particuliers", "54 755 €", "91", "42,3%"),
        ("Leroy Merlin Aubagne", "16 285 €", "29", "12,6%"),
        ("Leroy Merlin Toulon", "14 820 €", "33", "11,5%"),
        ("Leroy Merlin Cabriès", "13 945 €", "32", "10,8%"),
        ("Leroy Merlin La Valette", "10 755 €", "20", "8,3%"),
        ("Castorama", "10 140 €", "19", "7,8%"),
        ("Leroy Merlin Martigues", "7 760 €", "16", "6,0%"),
        ("Distrilap", "847 €", "7", "0,7%"),
    ]

    page.insert_text((50, y), "Client", fontsize=9, fontname="helv", color=GRIS)
    page.insert_text((220, y), "CA HT", fontsize=9, fontname="helv", color=GRIS)
    page.insert_text((320, y), "Factures", fontsize=9, fontname="helv", color=GRIS)
    page.insert_text((420, y), "Part CA", fontsize=9, fontname="helv", color=GRIS)
    y += 5
    page.draw_line((50, y), (500, y), color=GRIS, width=0.5)
    y += 15

    for client, ca, nb, pct in clients:
        page.insert_text((50, y), client, fontsize=9, fontname="helv", color=NOIR)
        page.insert_text((220, y), ca, fontsize=9, fontname="helv", color=NOIR)
        page.insert_text((320, y), nb, fontsize=9, fontname="helv", color=NOIR)
        page.insert_text((420, y), pct, fontsize=9, fontname="helv", color=NOIR)
        y += 16

    # Footer
    page.insert_text((50, 800), f"Rapport généré le {datetime.now().strftime('%d/%m/%Y')}",
                     fontsize=8, fontname="helv", color=GRIS)
    page.insert_text((450, 800), "Page 2/3", fontsize=8, fontname="helv", color=GRIS)

    # ===== PAGE 3 : KPIs et Recommandations =====
    page = doc.new_page(width=595, height=842)

    # Titre
    page.insert_text((50, 50), "INDICATEURS & RECOMMANDATIONS", fontsize=18, fontname="helv", color=BLEU)
    page.draw_line((50, 60), (545, 60), color=BLEU, width=1)

    # Zone d'intervention
    y = 90
    page.insert_text((50, y), "Zone d'Intervention", fontsize=12, fontname="helv", color=BLEU)
    y += 25

    zones = [
        ("Distance minimum", "3 km"),
        ("Distance maximum", "181 km (Nice)"),
        ("Distance moyenne", "51 km"),
        ("Total km/an (A/R)", "~42 500 km"),
        ("Durée moyenne trajet", "40 min"),
    ]

    for label, value in zones:
        page.insert_text((50, y), f"{label}:", fontsize=10, fontname="helv", color=GRIS)
        page.insert_text((200, y), value, fontsize=10, fontname="helv", color=NOIR)
        y += 18

    # KPIs
    y += 20
    page.insert_text((50, y), "Indicateurs de Performance", fontsize=12, fontname="helv", color=BLEU)
    y += 25

    kpis = [
        ("Chantiers par an", "~310"),
        ("Chantiers par mois", "26"),
        ("Chantiers par jour ouvré", "1,2"),
        ("CA par jour ouvré", "~590 €"),
        ("CA par km parcouru", "~3 €"),
    ]

    for label, value in kpis:
        page.insert_text((50, y), f"{label}:", fontsize=10, fontname="helv", color=GRIS)
        page.insert_text((200, y), value, fontsize=10, fontname="helv", color=NOIR)
        y += 18

    # Distribution des montants
    y += 20
    page.insert_text((50, y), "Distribution des Montants", fontsize=12, fontname="helv", color=BLEU)
    y += 25

    tranches = [
        ("0 - 200 €", "26", "10,5%"),
        ("200 - 400 €", "37", "15,0%"),
        ("400 - 600 €", "92", "37,2%"),
        ("600 - 800 €", "59", "23,9%"),
        ("800 - 1000 €", "33", "13,4%"),
    ]

    for tranche, nb, pct in tranches:
        page.insert_text((50, y), tranche, fontsize=10, fontname="helv", color=NOIR)
        page.insert_text((150, y), nb, fontsize=10, fontname="helv", color=NOIR)
        page.insert_text((200, y), pct, fontsize=10, fontname="helv", color=NOIR)
        # Barre graphique
        pct_val = float(pct.replace(',', '.').replace('%', ''))
        bar_width = pct_val * 3
        page.draw_rect(fitz.Rect(250, y - 8, 250 + bar_width, y + 2), color=BLEU, fill=BLEU)
        y += 18

    # Points forts
    y += 30
    page.insert_text((50, y), "Points Forts", fontsize=12, fontname="helv", color=(0, 0.5, 0))
    y += 20

    points_forts = [
        "Expertise reconnue (certification RGE)",
        "Partenariats solides avec les GSB",
        "Activité régulière et croissante (+23% au T4)",
        "Zone de chalandise étendue",
        "Mix clientèle équilibré B2B/Particuliers",
    ]

    for point in points_forts:
        page.insert_text((50, y), f"✓ {point}", fontsize=9, fontname="helv", color=NOIR)
        y += 15

    # Points d'attention
    y += 15
    page.insert_text((50, y), "Points d'Attention", fontsize=12, fontname="helv", color=(0.8, 0.4, 0))
    y += 20

    points_attention = [
        "Dépendance Leroy Merlin : 49% du CA",
        "Saisonnalité : creux en août",
        "Coût kilométrique : 42 500 km/an",
    ]

    for point in points_attention:
        page.insert_text((50, y), f"⚠ {point}", fontsize=9, fontname="helv", color=NOIR)
        y += 15

    # Recommandations
    y += 15
    page.insert_text((50, y), "Recommandations", fontsize=12, fontname="helv", color=BLEU)
    y += 20

    recommandations = [
        "1. Diversifier le portefeuille clients (nouveaux partenaires)",
        "2. Optimiser les tournées pour réduire les kilomètres",
        "3. Développer la clientèle particuliers (meilleure marge)",
        "4. Anticiper le creux estival (maintenance, formation)",
    ]

    for reco in recommandations:
        page.insert_text((50, y), reco, fontsize=9, fontname="helv", color=NOIR)
        y += 15

    # Footer
    page.insert_text((50, 800), f"Rapport généré le {datetime.now().strftime('%d/%m/%Y')}",
                     fontsize=8, fontname="helv", color=GRIS)
    page.insert_text((450, 800), "Page 3/3", fontsize=8, fontname="helv", color=GRIS)

    # Sauvegarder le PDF
    doc.save(OUTPUT_PDF)
    doc.close()

    print(f"PDF généré : {OUTPUT_PDF}")

if __name__ == "__main__":
    create_pdf_report()
