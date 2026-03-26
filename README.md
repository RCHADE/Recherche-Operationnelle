# Recherche Opérationnelle — OR-Tools

Travaux pratiques de Recherche Opérationnelle réalisés à l'IMT Mines Alès (spécialité PRISM).
Tous les modèles sont implémentés en Python avec le solveur OR-Tools de Google.

## Contenu

### 01 — Planification de Production (Programmation Linéaire)
Modélisation et résolution de problèmes de planification industrielle avec le solveur GLOP (nombres réels) et CBC (nombres entiers).

- **tp2_melange_acier.py** — Optimisation du mélange de matières premières pour la fabrication d'acier. Minimisation du coût sous contraintes de composition (carbone, cuivre, manganèse). Interface Excel.
- **tp2_planification_production.py** — Planification de la production d'aliments pour bétail. Minimisation du coût d'achat sous contraintes nutritionnelles. Interface Excel.

### 02 — Optimisation Combinatoire : TSP & Bin Packing
Résolution de problèmes d'optimisation combinatoire en nombres entiers avec le solveur CBC et Branch & Bound.

- **tp4_binpacking.py** — Chargement de flottes d'avions de fret (Bin Packing), découpe de bobines avec minimisation du nombre de bobines utilisées et de la perte matière (Cutting Stock).
- **tp5_tsp.py** — Tournée du voyageur de commerce (TSP) sur les 13 capitales de wilayas de Mauritanie (solution optimale : 3319 km) et sur les 48 moughataas. Élimination des sous-tournées par contraintes MTZ.

## Technologies
- Python 3
- [OR-Tools](https://developers.google.com/optimization) (Google) — solveurs GLOP et CBC
- openpyxl — interface Excel

## Structure
```
.
├── 01_Planification-Production-Lineaire/
│   ├── documents/          # Sujets PDF et fichiers Excel
│   ├── tp2_melange_acier.py
│   └── tp2_planification_production.py
├── 02_Optimisation-Combinatoire-TSP-BinPacking/
│   ├── documents/          # Sujets PDF et fichiers Excel
│   ├── tp4_binpacking.py
│   └── tp5_tsp.py
└── README.md
```