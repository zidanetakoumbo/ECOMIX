# Modèle financier solaire — feuille « Hypothèses »

Application statique HTML/CSS/JS (sans dépendance, sans build) qui reproduit la feuille
« Hypothèses » de `Modele_Financier_Solaire_3_Scenarios_GP V4.xlsx`.

## Structure

| Fichier | Rôle |
|---|---|
| `index.html` | Page et barre d'outils |
| `css/styles.css` | Thème Excel (couleurs centralisées dans `:root`) |
| `js/config.js` | Lignes de la feuille : libellés, cellules, valeurs par défaut, formats, commentaires |
| `js/formulas.js` | Formules Excel transcrites 1:1 (C9, C10, C30, G30, D38, D40, D45, D55, C75:G79) |
| `js/format.js` | Formats numériques Excel (fr-FR) et lecture des saisies |
| `js/app.js` | Rendu, recalcul, sauvegarde locale (localStorage), export JSON |

- Modifier une valeur par défaut ou un libellé : `js/config.js` (`def`, `label`, `note`).
- Ajouter une formule : une entrée dans `HYPO.FORMULAS` (`js/formulas.js`), dans l'ordre des dépendances.
- Les autres onglets pourront lire les valeurs via `HYPO.getValues()` ou l'événement `hypotheses:change`.

## Lancer en local

```bash
npx serve -l 5173 .
```

## Mise en production

Déposer le dossier tel quel sur n'importe quel hébergement statique (Vercel, Netlify, IIS, Nginx, Apache…).
