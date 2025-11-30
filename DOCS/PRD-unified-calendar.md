# PRD : Calendrier Unifié Multi-Départements

## Problématique

L'utilisateur dispose de 19 feuilles Google Sheets (une par département) contenant des événements avec les colonnes : Date, Service, Évènement, Sur site CCFHK?, Sur calendrier excel? (~300 événements au total). L'approche basée sur les formules (FILTER) a atteint les limites de calcul de Google Sheets, causant l'échec du calendrier après ~56 événements affichés.

## Exigences

- **Grille calendrier visuelle** (vue mensuelle traditionnelle) dans une nouvelle feuille
- **Filtres** : Année (cycle août-juillet) + Période + Service + Sur calendrier + Sur site
- **Année académique/liturgique** : Cycle août à juillet (ex: 2025-2026 = août 2025 - juillet 2026)
- **Maximum 8 événements par jour** affichés
- **Codage couleur par département** via emojis colorés
- **Mise à jour automatique** lors de modifications des données source
- **Accès partagé équipe** (Google Workspace Nonprofit)
- **Mois empilés verticalement** lors de l'affichage "Tout"
- **Lecture seule** (éditer dans les feuilles sources via hyperliens)

## Solution : Google Apps Script

### Pourquoi Apps Script (pas les formules)

Les approches basées sur les formules ne sont **pas viables** pour ce cas :

- 248 cellules × 19 feuilles × 300 lignes = ~1.4M évaluations de formules
- Dépasse les limites de calcul de Google Sheets
- Aucune optimisation possible dans les contraintes des formules

Avantages Apps Script :

- Lit toutes les 19 feuilles en ~0.5s (opération en masse)
- Filtre/groupe les événements en mémoire (~0.01s)
- Écrit la grille calendrier complète en ~0.5s
- **Exécution totale : ~1-2 secondes**

### Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                      Calendrier Unifié                               │
├─────────────────────────────────────────────────────────────────────┤
│ Lignes 1-3 : Contrôles de filtres                                   │
│   A1: "Année:"           | B1: [Dropdown: 2025-2026]                │
│   A2: "Période:"         | B2: [Dropdown: Actuel + À venir]         │
│   A3: "Service:"         | B3: [Dropdown: Tous]                     │
│   C1: "Sur calendrier:"  | D1: [Dropdown: Oui]                      │
│   C2: "Sur site:"        | D2: [Dropdown: Tous]                     │
│   E1: "Filtres par défaut:" | F1: [Checkbox]                        │
│   E2: "Mettre à jour:"      | F2: [Checkbox]                        │
│   E3: "Mis à jour: [timestamp]"                                     │
├─────────────────────────────────────────────────────────────────────┤
│ Ligne 4 : (vide - séparateur gelé)                                  │
├─────────────────────────────────────────────────────────────────────┤
│ Ligne 5+ : Grille calendrier                                        │
│   - En-tête du mois (fusionné, gras)                                │
│   - En-têtes jours : Lun|Mar|Mer|Jeu|Ven|Sam|Dim                   │
│   - Lignes de semaines (5-6 par mois)                               │
│   - Chaque cellule : Numéro du jour + jusqu'à 8 événements          │
│   - Emoji coloré par département + hyperlien vers source            │
└─────────────────────────────────────────────────────────────────────┘
```

### Fonctions principales

| Fonction | Objectif |
|----------|----------|
| `getAllEvents()` | Lire toutes les feuilles, retourner tableau d'événements normalisé |
| `getFilteredEvents()` | Appliquer filtres année + période + service + sur calendrier + sur site |
| `generateCalendarGrid()` | Construire grilles mensuelles, empiler verticalement |
| `renderCalendar()` | Orchestrer : lire filtres → récupérer données → générer → écrire |
| `applyFormatting()` | Emojis, polices, tailles cellules, fusions, hyperliens |
| `setupCalendarSheet()` | Créer feuille, dropdowns, rendu initial |
| `installCalendar()` | Installation en un clic |

### Stratégie de mise à jour

- **Mise à jour automatique** : Déclenchée par `onEdit()` lors de modification des feuilles départements
- **Mise à jour manuelle** : Case à cocher "Mettre à jour" ou menu "Calendrier > Mettre à jour"
- **Rafraîchissement auto** : Trigger toutes les 5 minutes (backup)
- **Horodatage** : Affiche "Mis à jour: [date/heure]" sous la case Mettre à jour
- **Message de chargement** : "Mise à jour en cours..." pendant le traitement

### Codage couleur par emojis

- 19 emojis colorés distincts assignés aux départements alphabétiquement
- Palette : 🔴 🟡 🟢 🔵 🟣 🟠 🩵 💚 💛 🌕 🧡 💜 💙 🩷 🩶 💚 🟠 ❤️ 🩷
- Préfixes spéciaux selon statut "Sur calendrier excel?" :
  - **Oui** : `🔴` (emoji département seul)
  - **Vide/Blanc** : `❓🔴` (point d'interrogation + emoji département)
  - **Non** : `🙈🔴` (singe qui se cache + emoji département)

### Options de filtres

**Année (B1) :**

- "2024-2025" (août 2024 - juillet 2025)
- "2025-2026" (août 2025 - juillet 2026) ← **Défaut**
- "2026-2027" (août 2026 - juillet 2027)

**Période (B2) :**

- "Tout" (tous les mois de l'année sélectionnée)
- "Actuel + À venir" (mois actuel + futurs) ← **Défaut**
- Mois individuels : "août", "septembre", ... "juillet"

**Service (B3) :**

- "Tous" ← **Défaut**
- Chaque nom de département (19 options)

**Sur calendrier (D1) :**

- "Tous"
- "Oui" ← **Défaut**
- "Non"

**Sur site (D2) :**

- "Tous" ← **Défaut**
- "Oui"
- "Non"

### Colonnes sources requises

Chaque feuille département doit avoir :

| Colonne | Requis | Description |
|---------|--------|-------------|
| Date | Oui | Date de l'événement |
| Service | Oui | Type de service |
| Évènement | Oui | Nom de l'événement (accepte variations d'accents) |
| Sur site CCFHK? | Non | Oui/Non - événement sur site |
| Sur calendrier excel? | Non | Oui/Non - à inclure dans calendrier Excel |

## Navigation vers la source

Chaque événement dans le calendrier est un **hyperlien** vers sa feuille source :

- Cliquer sur un événement ouvre la feuille département correspondante
- Fonctionne sur desktop et mobile
- Texte stylé en bleu souligné

## Fonctionnalités

### Mise en évidence "Aujourd'hui"

- La cellule du jour actuel est mise en évidence avec :
  - Fond bleu clair (#BBDEFB)
  - Bordure bleue (#1976D2)

### Semaine commençant le lundi

- Calendrier français : Lun | Mar | Mer | Jeu | Ven | Sam | Dim

### Actions via cases à cocher

- **Filtres par défaut** : Réinitialise tous les filtres aux valeurs par défaut
- **Mettre à jour** : Force une mise à jour immédiate du calendrier

## Notes techniques

- **Temps d'exécution** : <2 secondes pour 300 événements
- **Quota trigger** : 6 min/jour pour triggers basés sur le temps
- **Limite cellules** : ~120 lignes pour vue année complète
- **Locale française** : `fr-FR` pour noms de mois, abréviations jours

## Histoires utilisateur

1. **En tant que membre d'équipe**, je veux voir tous les événements départementaux dans une vue calendrier unifiée pour comprendre ce qui se passe dans l'organisation.

2. **En tant que coordinateur**, je veux filtrer par département pour me concentrer sur les événements pertinents à mon équipe.

3. **En tant que planificateur**, je veux filtrer par année et période pour voir uniquement les événements de l'année académique actuelle.

4. **En tant que responsable d'événement**, je veux éditer les événements dans ma feuille département et voir les changements reflétés immédiatement dans le calendrier unifié.

5. **En tant qu'utilisateur mobile**, je veux que le calendrier soit consultable sur mon téléphone via l'app Google Sheets.

6. **En tant qu'administrateur**, je veux identifier rapidement les événements sans statut "Sur calendrier excel" défini grâce à l'emoji ❓.

## Hors périmètre

- Édition directe depuis la vue calendrier (utiliser hyperliens pour naviguer vers la source)
- Intégration Google Calendar (pourrait être ajouté ultérieurement)
- Export PDF
- Interface personnalisation des couleurs

## Critères de succès

- [x] Tous les ~300 événements des 19 feuilles affichés correctement
- [x] Filtre par année fonctionne (cycle académique août-juillet)
- [x] Filtre par période fonctionne (Tout, Mois, Actuel+À venir)
- [x] Filtre par département fonctionne (Tous, département spécifique)
- [x] Filtre "Sur calendrier" fonctionne (Tous/Oui/Non)
- [x] Filtre "Sur site" fonctionne (Tous/Oui/Non)
- [x] Calendrier se rafraîchit en <3 secondes
- [x] Mise à jour automatique lors de modifications des données
- [x] Emojis colorés visibles pour identifier les départements
- [x] Hyperliens sur événements naviguent vers feuille source
- [x] Interface en français (noms de mois, jours)
- [x] Mise en évidence du jour actuel
- [x] Semaine commençant le lundi
- [x] Indicateurs visuels pour statut "Sur calendrier excel" (❓ pour vide, 🙈 pour Non)
