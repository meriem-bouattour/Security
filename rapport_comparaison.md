
# Rapport de Comparaison des Fichiers Excel
**Date de génération :** 2026-03-04 07:59  
**Fichier 1 (IDS) :** `security_ids.xlsx`  
**Fichier 2 (REFPROD) :** `security_refprod.xlsx`  
---

## 1. Résumé Exécutif
- **Lignes dans security_ids :** 13
- **Lignes dans security_refprod :** 13
- **Lignes communes (même ID_INTERNE) :** 10
- **Lignes uniques à security_ids :** 3
- **Lignes uniques à security_refprod :** 3
- **Colonnes dans security_ids :** 10
- **Colonnes dans security_refprod :** 10
- **Colonnes communes :** 4
- **Colonnes exclusives à security_ids :** 6
- **Colonnes exclusives à security_refprod :** 6
- **Colonnes communes avec divergences de valeur :** 1

## 2. Structure des Fichiers

### 2.1 Colonnes de security_ids
| # | Colonne | Présente dans REFPROD |
|---|---------|----------------------|
| 1 | `CLASSE_ACTIF` | ❌ Non |
| 2 | `CUSIP` | ❌ Non |
| 3 | `DATE_CREATION` | ❌ Non |
| 4 | `ID_INTERNE` | ✅ Oui |
| 5 | `ISIN` | ✅ Oui |
| 6 | `LIBELLE` | ❌ Non |
| 7 | `MARCHE` | ✅ Oui |
| 8 | `RIC` | ❌ Non |
| 9 | `SEDOL` | ❌ Non |
| 10 | `STATUT` | ✅ Oui |

### 2.2 Colonnes de security_refprod
| # | Colonne | Présente dans IDS |
|---|---------|-------------------|
| 1 | `DATE_MISE_A_JOUR` | ❌ Non |
| 2 | `DEVISE` | ❌ Non |
| 3 | `EMETTEUR` | ❌ Non |
| 4 | `ID_INTERNE` | ✅ Oui |
| 5 | `ISIN` | ✅ Oui |
| 6 | `MARCHE` | ✅ Oui |
| 7 | `NOTATION` | ❌ Non |
| 8 | `PAYS` | ❌ Non |
| 9 | `STATUT` | ✅ Oui |
| 10 | `TYPE_PRODUIT` | ❌ Non |

### 2.3 Colonnes exclusives
**Uniquement dans security_ids :** `CLASSE_ACTIF`, `CUSIP`, `DATE_CREATION`, `LIBELLE`, `RIC`, `SEDOL`
**Uniquement dans security_refprod :** `DATE_MISE_A_JOUR`, `DEVISE`, `EMETTEUR`, `NOTATION`, `PAYS`, `TYPE_PRODUIT`

## 3. Analyse des Lignes (clé : ID_INTERNE)

### 3.1 Lignes communes
10 identifiants présents dans les deux fichiers :

`SEC001`, `SEC002`, `SEC003`, `SEC004`, `SEC005`, `SEC006`, `SEC007`, `SEC008`, `SEC009`, `SEC010`

### 3.2 Lignes uniques à security_ids
3 ligne(s) absente(s) de security_refprod :

`SEC011`, `SEC012`, `SEC013`

### 3.3 Lignes uniques à security_refprod
3 ligne(s) absente(s) de security_ids :

`SEC014`, `SEC015`, `SEC016`

## 4. Cohérence des Codes ISIN
- ISIN communs aux deux fichiers : **10**
- ISIN uniquement dans security_ids : **3**
- ISIN uniquement dans security_refprod : **3**

### ISIN présents uniquement dans security_ids
- `FR0010208488`
- `GB00B03MLX29`
- `US2546871060`

### ISIN présents uniquement dans security_refprod
- `DE0005552004`
- `FR0013506730`
- `US38141G1040`

## 5. Divergences de Valeur sur les Colonnes Communes
1 colonne(s) présentent des différences de valeur :

### Colonne : `STATUT`
| ID_INTERNE | Valeur dans IDS | Valeur dans REFPROD |
|------------|-----------------|---------------------|
| `SEC007` | INACTIF | ACTIF |

## 6. Statistiques par Colonne Commune
| Colonne | Valeurs nulles IDS | Valeurs nulles REFPROD | Valeurs uniques IDS | Valeurs uniques REFPROD |
|---------|-------------------|------------------------|---------------------|-------------------------|
| `ISIN` | 0 | 0 | 13 | 13 |
| `MARCHE` | 0 | 0 | 5 | 5 |
| `STATUT` | 0 | 0 | 2 | 1 |

## 7. Conclusions et Points de Divergence
---

### Points de Similitude
- Les deux fichiers utilisent `ID_INTERNE` et `ISIN` comme identifiants clés.
- 10 valeurs d'`ID_INTERNE` sont partagées entre les deux fichiers.
- 4 colonne(s) sont communes : `ID_INTERNE`, `ISIN`, `MARCHE`, `STATUT`.
- La colonne `MARCHE` et `STATUT` apparaissent dans les deux fichiers avec des valeurs globalement cohérentes.

### Points de Divergence
- **Colonnes IDS uniquement :** `CLASSE_ACTIF`, `CUSIP`, `DATE_CREATION`, `LIBELLE`, `RIC`, `SEDOL` — ces colonnes d'identification ne sont pas portées par REFPROD.
- **Colonnes REFPROD uniquement :** `DATE_MISE_A_JOUR`, `DEVISE`, `EMETTEUR`, `NOTATION`, `PAYS`, `TYPE_PRODUIT` — données de référence produit absentes de IDS.
- **3 titre(s) présent(s) dans IDS mais absents de REFPROD** : SEC011, SEC012, SEC013. Ces titres n'ont pas encore de fiche produit de référence.
- **3 titre(s) présent(s) dans REFPROD mais absents de IDS** : SEC014, SEC015, SEC016. Ces titres n'ont pas d'identifiants de marché associés.
- **Divergences de valeur** sur 1 colonne(s) pour les lignes communes : `STATUT`.

### Recommandations
1. **Aligner les périmètres** : les titres SEC011, SEC012, SEC013 (IDS) et SEC014, SEC015, SEC016 (REFPROD) doivent être réconciliés.
2. **Compléter les fiches REFPROD** pour les titres présents uniquement dans IDS.
3. **Enrichir IDS** avec les données d'identification manquantes pour les titres REFPROD.
4. **Vérifier les divergences** de valeur colonne par colonne, notamment pour `STATUT`.
5. **Mettre en place un processus de synchronisation** régulier entre les deux référentiels.
