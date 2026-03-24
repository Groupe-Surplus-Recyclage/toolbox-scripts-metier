# divers_scripts

Répertoire de scripts utilitaires internes. Chaque script est autonome et répond à un besoin métier spécifique.

---

## Prérequis

- Python 3.10+
- Dépendances :

```bash
pip install openpyxl
```

---

## Scripts

### `immat_dc_pdf_lookup.py` — Recherche de PDF par immatriculation

Prend en entrée un fichier Excel avec les onglets **AUTO**, **MOTO** et **INDUSTRIE**, et pour chaque ligne recherche les PDFs correspondant à l'immatriculation (colonne D) dans les répertoires de scan réseau. Les résultats sont écrits directement dans le fichier Excel.

#### Utilisation

```bash
python immat_dc_pdf_lookup.py <chemin_vers_excel>
```

Exemple :

```bash
python immat_dc_pdf_lookup.py "C:\Users\gauthier.segonds\Desktop\divers_scripts\dc_to_check.xlsx"
```

#### Répertoires de scan par onglet

| Onglet     | Serveur    | Chemin                                              |
|------------|------------|-----------------------------------------------------|
| AUTO       | srvfile03  | `\\srvfile03\FICHIERS-ERP\ATEMO-2021\scan`          |
| INDUSTRIE  | srvfile03  | `\\srvfile03\FICHIERS-ERP\ATEMO-INDUSTRIE\Scan`     |
| MOTO       | srvfile06  | `\\srvfile06\DONNEES\Bases\Base_Surplus_Motos\Scan` |

#### Colonnes écrites dans l'Excel

| Colonne | Nom           | Contenu                                              |
|---------|---------------|------------------------------------------------------|
| J       | Nom du Fichier | Chemins complets des fichiers retenus (séparés par ` \| `) |
| K       | Cas           | Noms de fichiers seuls (sans chemin, séparés par ` \| `)   |
| L       | Nb Résultats  | Nombre de fichiers retenus (0 si aucun)              |

#### Logique de filtrage et priorité

Pour chaque immatriculation, la recherche suit cet ordre de priorité :

1. **`IMMAT_DC_RECEPISSE*`** — si présent, on ne garde que lui
2. **`IMMAT_DC*` ou `IMMAT_DOC_DECL_*`** — sinon, on garde ces fichiers
3. **Tous les fichiers trouvés** — fallback si aucun des patterns ci-dessus ne correspond

#### Exclusions

Les fichiers dont le nom contient l'un des patterns suivants sont systématiquement écartés (insensible à la casse) :

```
_be, " be", csa, certcession, cession, perte, achat, fiv,
_re*, vies, grais, lettre, offre, rapport,
photos, cmr, affaires personnelles, _pv, " pv", mad, frais,
" da", _da, mise en demeure, " ci", _ci, bon, positaire,
darva, assurance, facture, vroom, document, attestation,
confirmation, avis, fiche, enlevement, lar, cni
```

> **Exception** : `_re` n'exclut pas les fichiers contenant `_recepisse`.

#### Fonctionnement interne

Le script construit un **index en mémoire** de tous les PDFs du répertoire de scan en une seule passe réseau, puis effectue les recherches localement dans cet index. Cela évite de multiplier les accès réseau (un seul `os.walk` pour l'ensemble des lignes de l'onglet).

---

*Ce répertoire a vocation à accueillir d'autres scripts utilitaires. Ajouter une section par script en suivant le même format.*
