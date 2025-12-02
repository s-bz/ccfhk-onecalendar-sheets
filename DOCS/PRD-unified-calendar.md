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

### Couleurs d'arrière-plan des cellules

Le calendrier utilise des couleurs d'arrière-plan pour identifier visuellement les weekends et les jours fériés :

| Type de jour | Couleur | Code |
|--------------|---------|------|
| **Samedi** | Bleu clair | #E3F2FD |
| **Dimanche** | Orange clair | #FFE0B2 |
| **Vacances scolaires LFI** (jours de semaine) | Violet clair | #E1BEE7 |
| **Jours fériés Hong Kong** | Rouge clair | #FFCDD2 |
| **Jours normaux** | Blanc | #FFFFFF |

**Priorité des couleurs :**
1. Jours fériés publics → rouge clair
2. Vacances scolaires (lundi-vendredi) → violet clair
3. Samedi → bleu clair
4. Dimanche → orange clair

**Labels dans les cellules :**
- Jours fériés publics : "HK PH - [nom]" (ex: "HK PH - Noël")
- Vacances scolaires : "Vacances LFI"
- Weekend + vacances : couleur weekend + label "Vacances LFI"

### Calendrier des vacances et jours fériés

Le système inclut automatiquement les **vacances scolaires LFI/FIS** et les **jours fériés de Hong Kong** :

**Année 2025-2026 :**

Vacances scolaires :
- Vacances d'été : 1 août - 25 août 2025
- Vacances d'octobre : 24 octobre - 31 octobre 2025
- Vacances d'hiver : 22 décembre 2025 - 2 janvier 2026
- Vacances Nouvel An chinois : 16 février - 20 février 2026
- Vacances de Pâques : 30 mars - 10 avril 2026
- Vacances de printemps : 26 mai - 29 mai 2026

Jours fériés publics :
- 1 octobre 2025 : Fête nationale
- 7 octobre 2025 : Fête mi-automne
- 29 octobre 2025 : Chung Yeung
- 25-26 décembre 2025 : Noël / Boxing Day
- 1 janvier 2026 : Jour de l'an
- 17 février 2026 : Nouvel An chinois
- 3-7 avril 2026 : Pâques et Ching Ming
- 1 mai 2026 : Fête du travail
- 25 mai 2026 : Anniversaire de Bouddha
- 19 juin 2026 : Tuen Ng

**Note :** Pour mettre à jour les dates pour d'autres années académiques, modifier les objets `CONFIG.schoolHolidays` et `CONFIG.publicHolidays` dans Code.gs.

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
| Début | Oui* | Date/heure de début (nouveau format) |
| Fin | Non | Date/heure de fin (optionnel) |
| Date | Oui* | Date de l'événement (format legacy - rétrocompatible) |
| Service | Oui | Type de service |
| Évènement | Oui | Nom de l'événement (accepte variations d'accents) |
| Sur site CCFHK? | Non | Oui/Non - événement sur site |
| Sur calendrier excel? | Non | Oui/Non - à inclure dans calendrier Excel |

*Note: Soit "Début" (nouveau) soit "Date" (legacy) est requis. Le système détecte automatiquement le format.

### Support Date/Heure

Le calendrier supporte maintenant les événements avec heures et les événements multi-jours :

| Début | Fin | Résultat Google Calendar |
|-------|-----|--------------------------|
| `15/03/2025` | Vide | Événement journée entière |
| `15/03/2025` | `16/03/2025` | Événement multi-jours (journée entière) |
| `15/03/2025 14:00` | Vide | Événement avec heure (durée 1h par défaut) |
| `15/03/2025 14:00` | `15/03/2025 16:00` | Événement avec durée |
| `15/03/2025 14:00` | `17/03/2025 12:00` | Événement multi-jours avec heures |

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

---

## Synchronisation Google Calendar

### Problématique

Les utilisateurs veulent que les événements marqués "Sur site CCFHK? = Oui" soient automatiquement synchronisés vers un Google Calendar partagé pour accès mobile et intégration avec d'autres outils.

### Exigences

- **Filtre** : Seuls les événements avec `Sur site CCFHK? = Oui` sont synchronisés
- **Calendrier unique** : Un seul Google Calendar partagé pour tous les départements
- **Tags** : Nom du département dans la description au format `[[[NOM_DÉPARTEMENT]]]` pour filtrage
- **Sync incrémentale** : Suivi des IDs d'événements pour créer/modifier/supprimer uniquement les changements
- **Auto-sync** : Déclenché lors de modifications des feuilles (avec debounce de 30s)
- **Événements journée entière** : Pas d'heure dans les données source
- **Suppressions** : Si `Sur site CCFHK` devient Non/vide, l'événement est supprimé du Calendar

### Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                         Flux de données                              │
├─────────────────────────────────────────────────────────────────────┤
│                                                                      │
│  19 Feuilles Départements        Feuille _CalendarSync              │
│  ┌──────────────────┐           ┌─────────────────────────────┐     │
│  │ Date             │           │ eventHash | gcalEventId     │     │
│  │ Service          │ ────────> │ md5...    | abc123...       │     │
│  │ Evenement        │           │ md5...    | def456...       │     │
│  │ Sur site CCFHK?  │           └─────────────────────────────┘     │
│  └──────────────────┘                        │                       │
│         │                                    ▼                       │
│         │ Filtre:                ┌─────────────────────────────┐     │
│         │ Sur site = Oui         │    Google Calendar           │     │
│         └───────────────────────>│    [[[DÉPARTEMENT]]] tags    │     │
│                                  │    Événements journée        │     │
│                                  └─────────────────────────────┘     │
└─────────────────────────────────────────────────────────────────────┘
```

### Stratégie d'ID d'événement

Hash MD5 généré à partir des données (pas de modification des feuilles sources) :

```javascript
eventHash = MD5(département + "|" + date + "|" + service + "|" + événement)
```

### Format événement Calendar

```javascript
{
  title: "Service | Évènement",  // ou juste "Évènement" si pas de service
  description: "[[[NOM_DÉPARTEMENT]]]\n\nSource: CCFHK Events",
  start: { date: "2025-03-15" },  // Journée entière
  end: { date: "2025-03-16" }     // Journée entière = début + 1 jour
}
```

### Utilisation

1. **Sync manuelle** : Menu "Calendrier > Synchroniser Google Calendar"
2. **Sync automatique** : Déclenchée 30 secondes après modification d'une feuille département
3. **Filtrage dans Google Calendar** : Rechercher `[[[NOM_DÉPARTEMENT]]]` dans la description

### Fonctions principales

| Fonction | Objectif |
|----------|----------|
| `syncToGoogleCalendar()` | Orchestrateur principal |
| `getCalendar()` | Récupère le calendrier par ID |
| `getSyncTrackingData()` | Lit les données de suivi |
| `computeEventHash()` | Génère hash MD5 stable |
| `createCalendarEvent()` | Crée un événement |
| `updateCalendarEvent()` | Met à jour un événement |
| `deleteCalendarEvent()` | Supprime un événement |
| `writeSyncTracking()` | Écrit les données de suivi |
| `scheduleGoogleCalendarSync()` | Planifie sync avec debounce |

### Critères de succès - Google Calendar

- [x] Événements avec "Sur site CCFHK = Oui" apparaissent dans Google Calendar
- [x] Changements synchronisés dans les 60 secondes après modification
- [x] Événements supprimés/non marqués retirés du Calendar
- [x] Tags département `[[[NOM]]]` dans la description
- [x] Pas d'événements dupliqués
- [x] Sync survit aux erreurs API
- [x] Sync manuelle disponible via menu
- [x] Support événements avec heures (pas seulement journée entière)
- [x] Support événements multi-jours
- [x] Rétrocompatibilité avec colonne "Date" (legacy)
