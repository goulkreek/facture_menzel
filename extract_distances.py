#!/usr/bin/env python3
"""
Script d'extraction des adresses des devis et calcul des distances.
Génère un fichier CSV avec les trajets et kilomètres estimés.
"""

import os
import re
import csv
import math
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    print("Erreur: PyMuPDF non installé. Exécutez: pip3 install pymupdf")
    exit(1)

# Coordonnées du siège (La Cadière d'Azur, 83740)
SIEGE_LAT = 43.1936
SIEGE_LON = 5.7536
SIEGE_ADRESSE = "208 Chemin de la croix des signaux, 83740 La Cadière d'Azur"

# Base de données approximative des coordonnées GPS par code postal (région PACA)
# Format: code_postal: (latitude, longitude, nom_ville)
CODES_POSTAUX = {
    # Var (83)
    "83000": (43.1242, 5.9280, "Toulon"),
    "83100": (43.1242, 5.9280, "Toulon"),
    "83110": (43.1186, 5.8017, "Sanary-sur-Mer"),
    "83120": (43.2706, 6.5850, "Sainte-Maxime"),
    "83130": (43.1550, 6.0200, "La Garde"),
    "83140": (43.1708, 6.1347, "Six-Fours-les-Plages"),
    "83150": (43.1708, 5.7539, "Bandol"),
    "83160": (43.1375, 5.9825, "La Valette-du-Var"),
    "83170": (43.4089, 5.9267, "Tourves"),
    "83190": (43.3089, 5.7831, "Ollioules"),
    "83200": (43.1667, 6.1333, "Le Revest-les-Eaux"),
    "83210": (43.2050, 6.1097, "Solliès-Toucas"),
    "83220": (43.2056, 6.2417, "Le Pradet"),
    "83230": (43.0817, 6.2133, "Bormes-les-Mimosas"),
    "83240": (43.1456, 6.4631, "Cavalaire-sur-Mer"),
    "83250": (43.2006, 6.2706, "La Londe-les-Maures"),
    "83260": (43.2361, 6.2478, "La Crau"),
    "83270": (43.2417, 6.1550, "Saint-Cyr-sur-Mer"),
    "83280": (43.1000, 6.1333, "Le Lavandou"),
    "83300": (43.4264, 6.7675, "Draguignan"),
    "83310": (43.3119, 6.2689, "Cogolin"),
    "83320": (43.2333, 6.1000, "Carqueiranne"),
    "83330": (43.3331, 6.0506, "Le Beausset"),
    "83340": (43.2833, 6.0500, "Le Cannet-des-Maures"),
    "83350": (43.3167, 6.4000, "Ramatuelle"),
    "83360": (43.3336, 6.2378, "Gonfaron"),
    "83370": (43.2914, 6.0056, "Saint-Aygulf"),
    "83390": (43.3269, 6.2736, "Cuers"),
    "83400": (43.1236, 6.1378, "Hyères"),
    "83420": (43.1786, 6.1264, "La Croix-Valmer"),
    "83430": (43.2053, 6.0794, "Saint-Mandrier-sur-Mer"),
    "83440": (43.3644, 6.7478, "Fayence"),
    "83450": (43.2000, 6.0833, "Saint-Tropez"),
    "83460": (43.3833, 6.3000, "Les Arcs"),
    "83470": (43.1792, 6.0506, "Saint-Maximin-la-Sainte-Baume"),
    "83480": (43.2667, 6.8500, "Puget-sur-Argens"),
    "83490": (43.2911, 6.2536, "Le Muy"),
    "83500": (43.1728, 6.6289, "La Seyne-sur-Mer"),
    "83510": (43.2833, 6.6667, "Lorgues"),
    "83520": (43.2806, 6.3917, "Roquebrune-sur-Argens"),
    "83530": (43.2000, 6.1500, "Saint-Raphaël"),
    "83550": (43.3167, 6.1667, "Vidauban"),
    "83560": (43.3167, 6.0333, "Vinon-sur-Verdon"),
    "83570": (43.3333, 6.2167, "Carcès"),
    "83580": (43.2833, 6.3833, "Gassin"),
    "83590": (43.2833, 6.4667, "Giens"),
    "83600": (43.4244, 6.8431, "Fréjus"),
    "83610": (43.3167, 6.1333, "Collobrières"),
    "83620": (43.2833, 6.0833, "Flassans-sur-Issole"),
    "83630": (43.2667, 5.9667, "Aups"),
    "83640": (43.2583, 5.9250, "Plan-d'Aups-Sainte-Baume"),
    "83660": (43.3000, 6.2500, "Carnoules"),
    "83670": (43.3472, 6.2389, "Barjols"),
    "83680": (43.2167, 6.0500, "La Garde-Freinet"),
    "83690": (43.3333, 6.4167, "Salernes"),
    "83700": (43.2000, 6.2500, "Saint-Raphaël"),
    "83720": (43.3667, 6.0833, "Trans-en-Provence"),
    "83740": (43.1936, 5.7536, "La Cadière-d'Azur"),
    "83750": (43.2306, 5.9611, "La Seyne-sur-Mer"),
    "83780": (43.3167, 6.1500, "Flayosc"),
    "83790": (43.2333, 6.2333, "Pignans"),
    "83820": (43.1500, 6.1000, "Rayol-Canadel-sur-Mer"),
    "83830": (43.3000, 6.0167, "Bargemon"),
    "83840": (43.2500, 6.0833, "Comps-sur-Artuby"),
    "83850": (43.2167, 6.1000, "Nans-les-Pins"),
    "83860": (43.2000, 6.0333, "Néoules"),
    "83870": (43.2333, 6.1833, "Signes"),
    "83890": (43.1833, 5.9500, "Besse-sur-Issole"),
    "83910": (43.2833, 6.0167, "Pourrières"),
    "83920": (43.2667, 6.1167, "La Motte"),
    "83930": (43.2500, 6.1500, "Tourtour"),
    "83950": (43.1583, 6.4833, "La Môle"),
    "83980": (43.2833, 6.2500, "Le Lavandou"),
    "83990": (43.2167, 6.2667, "Saint-Tropez"),

    # Bouches-du-Rhône (13)
    "13001": (43.2965, 5.3698, "Marseille 1er"),
    "13002": (43.3036, 5.3656, "Marseille 2e"),
    "13003": (43.3106, 5.3789, "Marseille 3e"),
    "13004": (43.3067, 5.3989, "Marseille 4e"),
    "13005": (43.2917, 5.3958, "Marseille 5e"),
    "13006": (43.2889, 5.3806, "Marseille 6e"),
    "13007": (43.2833, 5.3667, "Marseille 7e"),
    "13008": (43.2500, 5.3833, "Marseille 8e"),
    "13009": (43.2417, 5.4333, "Marseille 9e"),
    "13010": (43.2833, 5.4167, "Marseille 10e"),
    "13011": (43.2917, 5.4500, "Marseille 11e"),
    "13012": (43.3083, 5.4333, "Marseille 12e"),
    "13013": (43.3250, 5.4167, "Marseille 13e"),
    "13014": (43.3333, 5.3833, "Marseille 14e"),
    "13015": (43.3583, 5.3583, "Marseille 15e"),
    "13016": (43.3583, 5.3167, "Marseille 16e"),
    "13100": (43.5263, 5.4454, "Aix-en-Provence"),
    "13104": (43.5833, 5.4000, "Arles"),
    "13105": (43.6139, 4.6275, "Arles"),
    "13106": (43.6139, 4.6275, "Arles"),
    "13109": (43.2917, 5.5417, "Simiane-Collongue"),
    "13111": (43.4417, 5.6083, "Coudoux"),
    "13112": (43.2500, 5.6000, "La Destrousse"),
    "13113": (43.2667, 5.4833, "Plan-de-Cuques"),
    "13114": (43.2833, 5.5333, "Puyloubier"),
    "13115": (43.4833, 4.9333, "Saint-Paul-lez-Durance"),
    "13117": (43.2667, 5.5000, "Lavéra"),
    "13118": (43.4167, 5.5000, "Istres"),
    "13119": (43.2833, 5.4833, "Saint-Savournin"),
    "13120": (43.3500, 5.4333, "Gardanne"),
    "13121": (43.6167, 4.8167, "Aurons"),
    "13122": (43.2500, 5.5000, "Ventabren"),
    "13124": (43.4667, 5.3833, "Peypin"),
    "13126": (43.5000, 5.3333, "Vauvenargues"),
    "13127": (43.4917, 5.0500, "Vitrolles"),
    "13128": (43.4167, 5.4000, "Plan-d'Orgon"),
    "13129": (43.6333, 4.8333, "Salin-de-Giraud"),
    "13130": (43.4500, 5.1833, "Berre-l'Étang"),
    "13140": (43.4000, 5.3167, "Miramas"),
    "13150": (43.3833, 5.1000, "Tarascon"),
    "13160": (43.4000, 5.1000, "Châteaurenard"),
    "13170": (43.4333, 5.2167, "Les Pennes-Mirabeau"),
    "13180": (43.3667, 5.2333, "Gignac-la-Nerthe"),
    "13190": (43.3500, 5.6667, "Allauch"),
    "13200": (43.4000, 4.9833, "Arles"),
    "13210": (43.3833, 4.8667, "Saint-Rémy-de-Provence"),
    "13220": (43.3167, 5.5000, "Châteauneuf-les-Martigues"),
    "13230": (43.4333, 4.9333, "Port-Saint-Louis-du-Rhône"),
    "13240": (43.4056, 5.3639, "Septèmes-les-Vallons"),
    "13250": (43.4500, 5.5000, "Saint-Chamas"),
    "13260": (43.2833, 5.0333, "Cassis"),
    "13270": (43.4000, 5.5333, "Fos-sur-Mer"),
    "13280": (43.4167, 5.3000, "Raphèle-lès-Arles"),
    "13290": (43.4500, 5.3500, "Aix-en-Provence"),
    "13300": (43.5333, 5.0667, "Salon-de-Provence"),
    "13310": (43.5000, 5.5333, "Saint-Martin-de-Crau"),
    "13320": (43.2667, 5.5500, "Bouc-Bel-Air"),
    "13330": (43.5500, 5.1833, "Pélissanne"),
    "13340": (43.4833, 5.5000, "Rognac"),
    "13350": (43.2833, 5.4167, "Charleval"),
    "13360": (43.3167, 5.2500, "Roquevaire"),
    "13370": (43.4500, 5.4500, "Mallemort"),
    "13380": (43.2833, 5.5000, "Plan-de-Cuques"),
    "13390": (43.3667, 5.5500, "Auriol"),
    "13400": (43.3833, 5.3000, "Aubagne"),
    "13410": (43.5167, 5.5500, "Lambesc"),
    "13420": (43.3667, 5.3833, "Gémenos"),
    "13430": (43.2500, 5.3833, "Eyguières"),
    "13440": (43.3000, 5.1000, "Cabannes"),
    "13450": (43.3833, 5.5500, "Grans"),
    "13460": (43.4167, 5.4000, "Les Saintes-Maries-de-la-Mer"),
    "13470": (43.4333, 5.5333, "Carnoux-en-Provence"),
    "13480": (43.2500, 5.6500, "Cabriès"),
    "13490": (43.2667, 5.5167, "Jouques"),
    "13500": (43.4058, 5.0475, "Martigues"),
    "13510": (43.2833, 5.2500, "Éguilles"),
    "13520": (43.2667, 5.5833, "Maussane-les-Alpilles"),
    "13530": (43.4833, 5.3833, "Trets"),
    "13540": (43.4500, 5.4333, "Puyricard"),
    "13550": (43.2833, 5.5333, "Noves"),
    "13560": (43.2667, 5.4833, "Sénas"),
    "13580": (43.2333, 5.2833, "La Fare-les-Oliviers"),
    "13590": (43.4167, 5.5333, "Meyreuil"),
    "13600": (43.1833, 5.6167, "La Ciotat"),
    "13610": (43.3333, 5.4833, "Le Puy-Sainte-Réparade"),
    "13620": (43.4333, 5.5000, "Carry-le-Rouet"),
    "13630": (43.3833, 5.5833, "Eyragues"),
    "13640": (43.2833, 5.5167, "La Roque-d'Anthéron"),
    "13650": (43.4333, 5.4333, "Meyrargues"),
    "13660": (43.3500, 5.5000, "Orgon"),
    "13670": (43.4500, 5.4500, "Saint-Andiol"),
    "13680": (43.5167, 5.4333, "Lançon-de-Provence"),
    "13690": (43.3667, 5.5333, "Graveson"),
    "13700": (43.4500, 5.4333, "Marignane"),
    "13710": (43.4833, 5.2667, "Fuveau"),
    "13720": (43.2833, 5.4833, "La Bouilladisse"),
    "13730": (43.4167, 5.5500, "Saint-Victoret"),
    "13740": (43.4833, 5.4833, "Le Rove"),
    "13750": (43.3167, 5.3333, "Plan-d'Orgon"),
    "13760": (43.5000, 5.5167, "Saint-Cannat"),
    "13770": (43.3333, 5.4500, "Venelles"),
    "13780": (43.3333, 5.4333, "Cuges-les-Pins"),
    "13790": (43.4500, 5.5000, "Rousset"),
    "13800": (43.4167, 4.9167, "Istres"),
    "13810": (43.3000, 5.2833, "Eygalières"),
    "13820": (43.3333, 5.5167, "Ensuès-la-Redonne"),
    "13821": (43.2833, 5.4500, "La Penne-sur-Huveaune"),
    "13830": (43.2833, 5.5333, "Roquefort-la-Bédoule"),
    "13840": (43.4167, 5.4833, "Rognes"),
    "13850": (43.3333, 5.4833, "Gréasque"),
    "13860": (43.3500, 5.4667, "Peyrolles-en-Provence"),
    "13870": (43.2833, 5.5000, "Rognonas"),
    "13880": (43.3333, 5.4500, "Velaux"),
    "13890": (43.3500, 5.5167, "Mouriès"),
    "13920": (43.3667, 5.4833, "Saint-Mitre-les-Remparts"),
    "13930": (43.2500, 5.5167, "Aureille"),
    "13950": (43.3333, 5.0833, "Cadolive"),
    "13960": (43.3833, 5.2833, "Sausset-les-Pins"),
    "13980": (43.3333, 5.3167, "Alleins"),
    "13990": (43.2833, 5.3333, "Fontvieille"),

    # Alpes-Maritimes (06)
    "06000": (43.7102, 7.2620, "Nice"),
    "06100": (43.7102, 7.2620, "Nice"),
    "06130": (43.5833, 7.1167, "Grasse"),
    "06140": (43.5667, 7.0833, "Vence"),
    "06150": (43.5333, 6.9333, "Cannes La Bocca"),
    "06160": (43.6500, 7.1500, "Antibes Juan-les-Pins"),
    "06200": (43.7102, 7.2620, "Nice"),
    "06210": (43.6167, 7.0333, "Mandelieu-la-Napoule"),
    "06220": (43.5833, 7.0167, "Golfe-Juan Vallauris"),
    "06230": (43.7500, 7.3833, "Villefranche-sur-Mer"),
    "06240": (43.6000, 7.0167, "Beausoleil"),
    "06250": (43.6000, 7.1167, "Mougins"),
    "06270": (43.5500, 7.0333, "Villeneuve-Loubet"),
    "06300": (43.7102, 7.2620, "Nice"),
    "06400": (43.5500, 6.9333, "Cannes"),
    "06410": (43.6000, 6.9833, "Biot"),
    "06480": (43.6167, 6.9333, "La Colle-sur-Loup"),
    "06500": (43.7667, 7.4500, "Menton"),
    "06530": (43.6333, 7.0000, "Peymeinade"),
    "06560": (43.5833, 7.0333, "Valbonne"),
    "06600": (43.6500, 7.1500, "Antibes"),
    "06610": (43.5833, 6.9833, "La Gaude"),
    "06620": (43.6833, 7.2667, "Le Bar-sur-Loup"),
    "06650": (43.5500, 6.9167, "Le Rouret"),
    "06700": (43.7167, 7.2167, "Saint-Laurent-du-Var"),
    "06730": (43.7000, 7.2167, "Saint-André-de-la-Roche"),
    "06740": (43.6000, 7.0167, "Châteauneuf-Grasse"),
    "06800": (43.6333, 7.1833, "Cagnes-sur-Mer"),
    "06810": (43.6167, 7.1333, "Auribeau-sur-Siagne"),
    "06830": (43.6500, 7.0167, "Gilette"),
    "06850": (43.6167, 7.1167, "Saint-Auban"),
    "06880": (43.7333, 7.3000, "Rocquebrune-Cap-Martin"),
    "06910": (43.6500, 6.9667, "Roquesteron"),
    "06950": (43.6833, 7.1333, "Falicon"),

    # Vaucluse (84)
    "84000": (43.9493, 4.8055, "Avignon"),
    "84100": (44.0556, 5.0469, "Orange"),
    "84110": (44.2333, 5.1500, "Vaison-la-Romaine"),
    "84120": (43.8333, 5.4000, "Pertuis"),
    "84130": (43.8833, 4.9500, "Le Pontet"),
    "84140": (43.9500, 5.0167, "Montfavet"),
    "84150": (43.9667, 5.0500, "Jonquières"),
    "84160": (43.7833, 5.2167, "Cadenet"),
    "84170": (43.8500, 4.8667, "Monteux"),
    "84190": (43.8333, 4.9667, "Beaumes-de-Venise"),
    "84200": (43.8000, 5.1167, "Carpentras"),
    "84210": (43.8667, 5.2667, "Pernes-les-Fontaines"),
    "84220": (43.8167, 5.2500, "Gordes"),
    "84230": (43.9333, 5.1000, "Châteauneuf-du-Pape"),
    "84240": (43.7500, 5.4833, "La Tour-d'Aigues"),
    "84250": (43.8333, 5.5000, "Le Thor"),
    "84260": (43.8500, 5.0167, "Sarrians"),
    "84270": (43.9500, 5.1333, "Vedène"),
    "84280": (43.9000, 5.2333, "Cavaillon"),
    "84290": (43.9333, 5.2500, "Cairanne"),
    "84300": (43.9167, 5.1167, "Cavaillon"),
    "84310": (43.8833, 5.3500, "Morières-lès-Avignon"),
    "84320": (43.9000, 5.0500, "Entraigues-sur-la-Sorgue"),
    "84330": (43.7833, 5.0833, "Caromb"),
    "84340": (43.8333, 5.2333, "Malaucène"),
    "84350": (43.8500, 5.0833, "Courthézon"),
    "84360": (43.8000, 5.3167, "Lauris"),
    "84370": (43.7667, 5.1667, "Bédarrides"),
    "84380": (43.8000, 5.2833, "Mazan"),
    "84390": (43.8833, 5.1833, "Sault"),
    "84400": (43.8667, 5.3333, "Apt"),
    "84410": (43.7500, 5.5167, "Bédoin"),
    "84420": (43.7833, 5.1667, "Piolenc"),
    "84430": (43.8500, 5.2500, "Mondragon"),
    "84440": (43.9000, 5.1833, "Robion"),
    "84450": (43.8167, 5.3000, "Saint-Saturnin-lès-Avignon"),
    "84460": (43.8333, 5.4167, "Cheval-Blanc"),
    "84470": (43.9167, 5.1000, "Châteauneuf-de-Gadagne"),
    "84480": (43.7833, 5.4000, "Bonnieux"),
    "84490": (43.9333, 5.0667, "Saint-Saturnin-lès-Apt"),
    "84500": (43.8333, 5.0500, "Bollène"),
    "84510": (43.7333, 5.1833, "Caumont-sur-Durance"),
    "84530": (43.8500, 5.1000, "Villelaure"),
    "84540": (43.8000, 5.3500, "Cheval-Blanc"),
    "84550": (43.8000, 5.0667, "Mornas"),
    "84560": (43.8500, 5.1833, "Ménerbes"),
    "84570": (43.8167, 5.1667, "Villes-sur-Auzon"),
    "84580": (43.8833, 5.2667, "Oppède"),
    "84600": (43.8333, 5.3000, "Valréas"),
    "84700": (43.8667, 5.2000, "Sorgues"),
    "84740": (43.8000, 5.4333, "Velleron"),
    "84750": (43.8833, 5.0333, "Viens"),
    "84800": (43.9206, 5.0111, "L'Isle-sur-la-Sorgue"),
    "84810": (43.8000, 5.2000, "Aubignan"),
    "84820": (43.8000, 5.2500, "Visan"),
    "84830": (43.7833, 5.2333, "Sérignan-du-Comtat"),
    "84840": (43.8167, 5.2000, "Lapalud"),
    "84850": (43.8333, 5.1667, "Travaillan"),
    "84860": (43.8167, 5.1333, "Caderousse"),
    "84870": (43.8333, 5.0833, "Loriol-du-Comtat"),
}

def haversine_distance(lat1, lon1, lat2, lon2):
    """
    Calcule la distance à vol d'oiseau entre deux points GPS (en km).
    Utilise la formule de Haversine.
    """
    R = 6371  # Rayon de la Terre en km

    lat1_rad = math.radians(lat1)
    lat2_rad = math.radians(lat2)
    delta_lat = math.radians(lat2 - lat1)
    delta_lon = math.radians(lon2 - lon1)

    a = math.sin(delta_lat / 2) ** 2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(delta_lon / 2) ** 2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

    return R * c

def extraire_code_postal(adresse):
    """Extrait le code postal d'une adresse."""
    match = re.search(r'\b(0[1-9]|[1-8][0-9]|9[0-5])\d{3}\b', adresse)
    if match:
        return match.group(0)
    return None

def nettoyer_texte(texte):
    """Nettoie le texte en supprimant les retours à la ligne superflus."""
    # Remplacer les retours à la ligne multiples par un seul
    texte = re.sub(r'\n+', '\n', texte)
    # Supprimer les espaces multiples
    texte = re.sub(r' +', ' ', texte)
    return texte.strip()

def extraire_infos_devis(texte):
    """
    Extrait les informations d'un devis à partir du texte.
    Retourne: (numero_devis, date, client, adresse_complete, code_postal, ville)
    """
    texte = nettoyer_texte(texte)

    # Numéro de devis
    match_devis = re.search(r'Devis\s*N[°o]?\s*:?\s*(\d+)', texte, re.IGNORECASE)
    numero_devis = match_devis.group(1) if match_devis else "N/A"

    # Date
    match_date = re.search(r'Date\s*:?\s*(\d{2}/\d{2}/\d{4})', texte)
    date = match_date.group(1) if match_date else "N/A"

    # Client et adresse - pattern amélioré
    # Pattern: "Client : Nom\nAdresse\nCode postal VILLE\nTEL"
    match_client = re.search(
        r'Client\s*:\s*([^\n]+)\n([^\n]+)\n(\d{5})\s+([A-Z][A-Z\s\-\']+?)(?:\n|TEL|$)',
        texte, re.IGNORECASE
    )

    if match_client:
        client = match_client.group(1).strip()
        rue = match_client.group(2).strip()
        code_postal = match_client.group(3)
        ville = match_client.group(4).strip()
        # Nettoyer la ville (supprimer TEL, espaces, etc.)
        ville = re.sub(r'\s*(TEL|REF|FACTURE).*$', '', ville, flags=re.IGNORECASE).strip()
        adresse_complete = f"{rue}, {code_postal} {ville}"
    else:
        # Essayer un autre pattern
        match_client2 = re.search(r'Client\s*:\s*([^\n]+)', texte, re.IGNORECASE)
        client = match_client2.group(1).strip() if match_client2 else "N/A"

        # Chercher le code postal et ville
        match_cp_ville = re.search(r'(\d{5})\s+([A-Z][A-Z\s\-\']+?)(?:\n|TEL|REF|FACTURE|$)', texte, re.IGNORECASE)
        if match_cp_ville:
            code_postal = match_cp_ville.group(1)
            ville = match_cp_ville.group(2).strip()
            ville = re.sub(r'\s*(TEL|REF|FACTURE).*$', '', ville, flags=re.IGNORECASE).strip()

            # Chercher la rue avant le code postal
            match_rue = re.search(rf'Client[^\n]*\n([^\n]+)\n[^\n]*{code_postal}', texte)
            if match_rue:
                rue = match_rue.group(1).strip()
            else:
                rue = ""
            adresse_complete = f"{rue}, {code_postal} {ville}".strip(", ")
        else:
            code_postal = None
            ville = "N/A"
            adresse_complete = "N/A"

    # Nettoyage final des données
    client = re.sub(r'\s+', ' ', client).strip()
    ville = re.sub(r'\s+', ' ', ville).strip()
    adresse_complete = re.sub(r'\s+', ' ', adresse_complete).strip()

    return numero_devis, date, client, adresse_complete, code_postal, ville

def calculer_distance(code_postal):
    """
    Calcule la distance à vol d'oiseau depuis le siège vers le code postal donné.
    Retourne la distance en km ou None si le code postal n'est pas trouvé.
    """
    if not code_postal:
        return None

    if code_postal in CODES_POSTAUX:
        lat, lon, _ = CODES_POSTAUX[code_postal]
        return round(haversine_distance(SIEGE_LAT, SIEGE_LON, lat, lon), 1)

    # Essayer de trouver un code postal proche (même préfixe)
    prefix = code_postal[:3]
    for cp, (lat, lon, _) in CODES_POSTAUX.items():
        if cp.startswith(prefix):
            return round(haversine_distance(SIEGE_LAT, SIEGE_LON, lat, lon), 1)

    return None

def traiter_pdf(chemin_pdf):
    """
    Traite un fichier PDF et extrait les informations du devis.
    """
    try:
        doc = fitz.open(chemin_pdf)
        texte = ""
        for page in doc:
            texte += page.get_text()
        doc.close()

        numero, date, client, adresse, code_postal, ville = extraire_infos_devis(texte)
        distance = calculer_distance(code_postal)

        return {
            'fichier': os.path.basename(chemin_pdf),
            'numero_devis': numero,
            'date': date,
            'client': client,
            'adresse_depart': SIEGE_ADRESSE,
            'adresse_destination': adresse,
            'code_postal': code_postal or "N/A",
            'ville': ville,
            'distance_km': distance if distance else "N/A"
        }
    except Exception as e:
        print(f"Erreur lors du traitement de {chemin_pdf}: {e}")
        return None

def main():
    """Fonction principale."""
    dossier_devis = Path("/Users/sroyer/Documents/Dev/maconnerieconcept/Factures_Menzel/DEVIS_2025")
    fichier_csv = dossier_devis.parent / "trajets_devis.csv"

    # Lister tous les PDFs
    pdfs = list(dossier_devis.glob("*.pdf"))
    print(f"Nombre de PDFs trouvés: {len(pdfs)}")

    resultats = []
    erreurs = 0

    for i, pdf in enumerate(pdfs, 1):
        if i % 50 == 0:
            print(f"Traitement: {i}/{len(pdfs)}")

        resultat = traiter_pdf(pdf)
        if resultat:
            resultats.append(resultat)
        else:
            erreurs += 1

    # Filtrer: garder seulement les lignes avec numéro de devis ET kilométrage valides
    resultats_valides = [
        r for r in resultats
        if r['numero_devis'] != "N/A"
        and r['numero_devis'].isdigit()
        and isinstance(r['distance_km'], (int, float))
    ]

    # Ajouter la colonne aller-retour
    for r in resultats_valides:
        r['distance_km_ar'] = round(r['distance_km'] * 2, 1)

    # Trier par numéro de devis
    resultats_valides.sort(key=lambda x: int(x['numero_devis']))

    # Écrire le CSV avec virgules pour les décimaux (format français)
    with open(fichier_csv, 'w', newline='', encoding='utf-8') as f:
        fieldnames = ['numero_devis', 'date', 'client', 'adresse_depart', 'adresse_destination', 'code_postal', 'ville', 'distance_km', 'distance_km_ar']
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
        writer.writeheader()

        for r in resultats_valides:
            row = {k: r[k] for k in fieldnames}
            # Convertir les distances en format français (virgule)
            row['distance_km'] = str(r['distance_km']).replace('.', ',')
            row['distance_km_ar'] = str(r['distance_km_ar']).replace('.', ',')
            writer.writerow(row)

    print(f"Lignes filtrées (avec N/A): {len(resultats) - len(resultats_valides)}")
    resultats = resultats_valides  # Pour les stats

    print(f"\n=== RÉSUMÉ ===")
    print(f"PDFs traités: {len(resultats)}")
    print(f"Erreurs: {erreurs}")
    print(f"Fichier CSV généré: {fichier_csv}")

    # Statistiques sur les distances
    distances = [r['distance_km'] for r in resultats if isinstance(r['distance_km'], (int, float))]
    if distances:
        print(f"\n=== STATISTIQUES DISTANCES ===")
        print(f"Distance minimale: {min(distances):.1f} km")
        print(f"Distance maximale: {max(distances):.1f} km")
        print(f"Distance moyenne: {sum(distances)/len(distances):.1f} km")
        print(f"Total kilomètres (aller simple): {sum(distances):.1f} km")
        print(f"Total kilomètres (aller-retour): {sum(distances)*2:.1f} km")

if __name__ == "__main__":
    main()
