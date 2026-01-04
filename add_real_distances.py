#!/usr/bin/env python3
"""
Script pour ajouter les distances routières réelles au CSV des trajets.
Utilise l'API Google Maps Distance Matrix.
"""

import csv
import time
import json
import urllib.request
import urllib.parse
from pathlib import Path
from typing import Dict, List, Tuple, Optional

# Configuration
API_KEY = "AIzaSyBEDwGB8NtGiNgWGXimJqTX_ketM7JVMKY"
SIEGE = "208 Chemin de la croix des signaux, 83740 La Cadière d'Azur, France"
BATCH_SIZE = 25  # Max destinations par requête API
DELAY_BETWEEN_BATCHES = 1  # Secondes entre les lots pour respecter les quotas

CSV_INPUT = Path("/Users/sroyer/Documents/Dev/maconnerieconcept/Factures_Menzel/trajets_devis.csv")
CSV_OUTPUT = Path("/Users/sroyer/Documents/Dev/maconnerieconcept/Factures_Menzel/trajets_devis_complet.csv")
CACHE_FILE = Path("/Users/sroyer/Documents/Dev/maconnerieconcept/Factures_Menzel/distances_cache.json")


def load_cache() -> Dict:
    """Charge le cache des distances depuis le fichier JSON."""
    if CACHE_FILE.exists():
        with open(CACHE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_cache(cache: Dict):
    """Sauvegarde le cache des distances."""
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def get_distances_batch(destinations: List[str]) -> Dict[str, Tuple[Optional[float], Optional[int]]]:
    """
    Récupère les distances routières pour un lot de destinations.
    Retourne un dict: {adresse: (distance_km, duree_minutes)}
    """
    results = {}

    # Construire l'URL de l'API
    base_url = "https://maps.googleapis.com/maps/api/distancematrix/json"

    params = {
        'origins': SIEGE,
        'destinations': '|'.join(destinations),
        'mode': 'driving',
        'language': 'fr',
        'key': API_KEY
    }

    url = f"{base_url}?{urllib.parse.urlencode(params)}"

    try:
        with urllib.request.urlopen(url, timeout=30) as response:
            data = json.loads(response.read().decode('utf-8'))

        if data['status'] != 'OK':
            print(f"Erreur API: {data['status']}")
            return results

        # Parser les résultats
        elements = data['rows'][0]['elements']
        destination_addresses = data['destination_addresses']

        for i, element in enumerate(elements):
            original_dest = destinations[i]

            if element['status'] == 'OK':
                distance_m = element['distance']['value']  # en mètres
                duration_s = element['duration']['value']  # en secondes

                distance_km = round(distance_m / 1000, 1)
                duration_min = round(duration_s / 60)

                results[original_dest] = (distance_km, duration_min)
            else:
                print(f"  Pas de route trouvée pour: {original_dest[:50]}... ({element['status']})")
                results[original_dest] = (None, None)

    except Exception as e:
        print(f"Erreur lors de l'appel API: {e}")

    return results


def main():
    """Fonction principale."""
    print("=== Ajout des distances routières réelles ===\n")

    # Charger le cache existant
    cache = load_cache()
    print(f"Cache chargé: {len(cache)} adresses en cache\n")

    # Lire le CSV source
    rows = []
    with open(CSV_INPUT, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter=';')
        rows = list(reader)

    print(f"CSV chargé: {len(rows)} lignes\n")

    # Extraire les adresses uniques non présentes dans le cache
    all_destinations = set()
    for row in rows:
        dest = row['adresse_destination']
        if dest and dest != 'N/A':
            all_destinations.add(dest)

    # Filtrer celles déjà en cache
    destinations_to_fetch = [d for d in all_destinations if d not in cache]

    print(f"Adresses uniques: {len(all_destinations)}")
    print(f"Déjà en cache: {len(all_destinations) - len(destinations_to_fetch)}")
    print(f"À récupérer: {len(destinations_to_fetch)}\n")

    # Récupérer les distances par lots
    if destinations_to_fetch:
        destinations_list = list(destinations_to_fetch)
        total_batches = (len(destinations_list) + BATCH_SIZE - 1) // BATCH_SIZE

        for i in range(0, len(destinations_list), BATCH_SIZE):
            batch = destinations_list[i:i + BATCH_SIZE]
            batch_num = i // BATCH_SIZE + 1

            print(f"Lot {batch_num}/{total_batches} ({len(batch)} adresses)...")

            batch_results = get_distances_batch(batch)

            # Ajouter au cache
            for dest, (dist, dur) in batch_results.items():
                cache[dest] = {'distance_km': dist, 'duree_min': dur}

            # Sauvegarder le cache après chaque lot
            save_cache(cache)

            # Attendre entre les lots
            if i + BATCH_SIZE < len(destinations_list):
                time.sleep(DELAY_BETWEEN_BATCHES)

        print(f"\nRécupération terminée. Cache mis à jour: {len(cache)} adresses\n")

    # Enrichir les données avec les distances réelles
    for row in rows:
        dest = row['adresse_destination']

        if dest in cache and cache[dest]['distance_km'] is not None:
            dist = cache[dest]['distance_km']
            dur = cache[dest]['duree_min']

            # Format français (virgule)
            row['distance_reelle_km'] = str(dist).replace('.', ',')
            row['distance_reelle_km_ar'] = str(round(dist * 2, 1)).replace('.', ',')
            row['duree_trajet_min'] = str(dur)
            row['duree_trajet_ar_min'] = str(dur * 2)
        else:
            row['distance_reelle_km'] = 'N/A'
            row['distance_reelle_km_ar'] = 'N/A'
            row['duree_trajet_min'] = 'N/A'
            row['duree_trajet_ar_min'] = 'N/A'

    # Écrire le nouveau CSV
    fieldnames = [
        'numero_devis', 'date', 'client', 'adresse_depart', 'adresse_destination',
        'code_postal', 'ville', 'distance_km', 'distance_km_ar',
        'distance_reelle_km', 'distance_reelle_km_ar', 'duree_trajet_min', 'duree_trajet_ar_min'
    ]

    with open(CSV_OUTPUT, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
        writer.writeheader()
        writer.writerows(rows)

    print(f"CSV généré: {CSV_OUTPUT}")

    # Statistiques
    distances_reelles = []
    for row in rows:
        try:
            dist = float(row['distance_reelle_km'].replace(',', '.'))
            distances_reelles.append(dist)
        except (ValueError, AttributeError):
            pass

    if distances_reelles:
        print(f"\n=== STATISTIQUES DISTANCES ROUTIÈRES ===")
        print(f"Trajets avec distance: {len(distances_reelles)}/{len(rows)}")
        print(f"Distance minimale: {min(distances_reelles):.1f} km")
        print(f"Distance maximale: {max(distances_reelles):.1f} km")
        print(f"Distance moyenne: {sum(distances_reelles)/len(distances_reelles):.1f} km")
        print(f"Total kilomètres (aller simple): {sum(distances_reelles):.1f} km")
        print(f"Total kilomètres (aller-retour): {sum(distances_reelles)*2:.1f} km")

        # Comparaison avec les distances à vol d'oiseau
        distances_vol = []
        for row in rows:
            try:
                dist = float(row['distance_km'].replace(',', '.'))
                distances_vol.append(dist)
            except (ValueError, AttributeError):
                pass

        if distances_vol and len(distances_vol) == len(distances_reelles):
            ratio_moyen = sum(distances_reelles) / sum(distances_vol)
            print(f"\n=== COMPARAISON ===")
            print(f"Ratio moyen route/vol d'oiseau: {ratio_moyen:.2f}x")
            print(f"Les distances réelles sont en moyenne {(ratio_moyen-1)*100:.0f}% plus longues")


if __name__ == "__main__":
    main()
