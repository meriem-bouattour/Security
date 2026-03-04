# Security

Comparaison des fichiers Excel `security_ids` et `security_refprod`.

## 📄 Rapport de comparaison

👉 **[rapport_comparaison.md](./rapport_comparaison.md)**

Le rapport détaille :
- Les différences et similitudes de structure (colonnes, lignes)
- Les lignes communes et uniques à chaque fichier (clé : `ID_INTERNE`)
- La cohérence des codes ISIN entre les deux fichiers
- Les divergences de valeur sur les colonnes communes
- Des statistiques par colonne (valeurs nulles, cardinalités)
- Un résumé global avec recommandations

## 📁 Structure du dépôt

```
Security/
├── data/
│   ├── security_ids.xlsx       # Fichier des identifiants de marché (13 titres, 10 colonnes)
│   └── security_refprod.xlsx   # Fichier de référence produit   (13 titres, 10 colonnes)
├── generate_and_compare.py     # Script Python : génère les Excel et produit le rapport
└── rapport_comparaison.md      # ← Rapport de comparaison détaillé
```

## ▶️ Régénérer le rapport

```bash
pip install pandas openpyxl xlsxwriter
python generate_and_compare.py
```