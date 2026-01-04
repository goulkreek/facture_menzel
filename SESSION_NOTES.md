# Session Notes - Configuration MCP Google Maps

## Ce qui a été fait

### 1. Extraction des trajets des devis (TERMINÉ)
- Script Python créé : `extract_distances.py`
- Fichier CSV généré : `trajets_devis.csv`
- **416 devis** traités avec distances à vol d'oiseau
- Format français (virgule pour décimaux)
- Colonnes : numero_devis, date, client, adresse_depart, adresse_destination, code_postal, ville, distance_km, distance_km_ar

### Statistiques extraites :
- Distance minimale : 2,5 km
- Distance maximale : 134,6 km (Nice)
- Distance moyenne : 34,2 km
- Total aller simple : 14 221 km
- Total aller-retour : 28 441 km

### 2. Installation gcloud (TERMINÉ)
- Google Cloud SDK installé via Homebrew
- Version : 550.0.0
- Chemin : `/opt/homebrew/share/google-cloud-sdk/bin/gcloud`

## Ce qui a été fait (suite)

### 3. Configuration Google Maps API (TERMINÉ)
- Connecté à gcloud avec sroyer@wizapp.fr
- Projet : wizappmanager
- API Distance Matrix activée
- Clé API créée : `AIzaSyBEDwGB8NtGiNgWGXimJqTX_ketM7JVMKY`

### 4. Configuration MCP Google Maps (TERMINÉ)
- Serveur MCP ajouté dans ~/.claude/claude_code_config.json
- **IMPORTANT : Redémarrer Claude Code pour activer le serveur MCP**

## Ce qui a été fait (session 2)

### 5. Distances routières réelles (TERMINÉ)
- Script créé : `add_real_distances.py`
- Utilise l'API Google Maps Distance Matrix
- Cache JSON pour éviter les appels API répétés : `distances_cache.json`
- Nouveau CSV généré : `trajets_devis_complet.csv`

### Nouvelles colonnes ajoutées :
- `distance_reelle_km` : distance routière aller (en voiture)
- `distance_reelle_km_ar` : distance routière aller-retour
- `duree_trajet_min` : durée du trajet aller (en minutes)
- `duree_trajet_ar_min` : durée du trajet aller-retour

### Statistiques distances routières :
- Distance minimale : 3,0 km
- Distance maximale : 180,9 km (Nice)
- Distance moyenne : 51,1 km
- Total aller simple : 21 255,6 km
- Total aller-retour : 42 511,2 km

### Comparaison vol d'oiseau vs route :
- Ratio moyen : 1,49x
- Les distances réelles sont en moyenne **49% plus longues** que les distances à vol d'oiseau

## Fichiers créés
- `extract_distances.py` - Extraction initiale (distances vol d'oiseau)
- `trajets_devis.csv` - CSV avec distances à vol d'oiseau
- `add_real_distances.py` - Ajout des distances routières via Google Maps API
- `trajets_devis_complet.csv` - **CSV final avec distances réelles**
- `distances_cache.json` - Cache des distances (379 adresses)
- `SESSION_NOTES.md` - Notes de session

## Adresse du siège (pour référence)
208 Chemin de la croix des signaux, 83740 La Cadière d'Azur
Coordonnées GPS : 43.1936° N, 5.7536° E
