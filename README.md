# Power BI – Excel → SQL Server → KPI → Dashboard (chargement quotidien automatique)

Projet d'exemple : des fichiers Excel déposés chaque jour dans un dossier daté sont chargés
automatiquement dans SQL Server, puis le dashboard Power BI est actualisé et affiche la tendance par jour.

```
inbox\2026-11-08\sales_2026-11-08.xlsx      (un dossier par jour, nom = AAAA-MM-JJ)
        │  tâche planifiée → run_daily.ps1 → load_daily.py
        │    • date déjà chargée (LoadLog = Success) → on ne fait rien
        │    • date nouvelle → tables stg_* → usp_LoadDay (incrémental) → LoadLog
        ▼
SQL Server « SalesKpiDW » : schéma en étoile + vues vw_MonthlyKpi, vw_DailyKpi, vw_CustomerRfm
        │  refresh_powerbi.ps1 (seulement si de nouvelles données) ou actualisation planifiée du service
        ▼
Power BI : modèle + mesures DAX (powerbi/measures.dax) + pages (powerbi/dashboard_layout.md)
```

## 1. Prérequis
| Outil | Remarque |
|---|---|
| SQL Server (Express suffit) + SSMS ou `sqlcmd` | serveur par défaut du projet : `localhost\SQLEXPRESS` |
| ODBC Driver 17 pour SQL Server | si vous avez le 18, remplacez `17` par `18` dans `CONN` |
| Python 3.9+ | |
| Power BI Desktop | gratuit (Microsoft Store) |

## 2. Installation (une seule fois)
Dans PowerShell, depuis le dossier du projet :
```powershell
# a) Paquets Python
pip install -r requirements.txt

# b) Base de données : tables, schéma en étoile, vues   (ATTENTION : 01 efface les tables existantes)
sqlcmd -S "localhost\SQLEXPRESS" -E -C -i sql\01_schema.sql

# c) Chargement quotidien : LoadLog, procédure usp_LoadDay, vue vw_DailyKpi   (non destructif)
sqlcmd -S "localhost\SQLEXPRESS" -E -C -i sql\02_daily_load.sql
```
Ou tout en une commande : `.\setup.ps1` (installe, crée la base et charge l'historique d'exemple).

Un autre serveur ? Modifiez `SERVER=` dans `CONN` (en haut de `load_excel_to_db.py` et `load_daily.py`)
et `-S` dans les commandes `sqlcmd`.

## 3. Chargement
### 3.1 Historique initial (une fois)
```powershell
python generate_sample_excel.py      # crée data\sales_data.xlsx (données d'exemple)
python load_excel_to_db.py           # ou : python load_excel_to_db.py "C:\chemin\fichier.xlsx"
```
### 3.2 Chargement quotidien
1. Déposez le fichier du jour dans un dossier nommé par la date : `inbox\2026-11-08\mon_fichier.xlsx`
   (feuille **Sales** obligatoire : `OrderID, OrderDate, CustomerID, ProductID, Quantity, DiscountPct` ;
   feuilles facultatives : `Customers`, `Products`, `Targets`).
2. Lancez :
   ```powershell
   python load_daily.py
   ```
   - date déjà chargée → `already loaded -> skipped`, rien n'est fait
   - date nouvelle → chargée dans la base et inscrite dans `LoadLog`

   Codes de sortie : `0` nouvelles données chargées · `2` rien de nouveau · `1` échec (voir `LoadLog` et `logs\`).
3. Pour tester sans vos données : `python generate_daily_files.py 2026-11-01 7` crée 7 dossiers de test.

### 3.3 Automatiser (tous les jours)
```powershell
.\register_daily_task.ps1 -Time 06:30      # crée la tâche Windows « SalesKpi_DailyLoad »
```
`run_daily.ps1` charge les nouveaux dossiers, écrit un journal dans `logs\`, puis déclenche l'actualisation
Power BI **uniquement** s'il y a eu de nouvelles données.

### 3.4 Vérifier
```powershell
sqlcmd -S "localhost\SQLEXPRESS" -E -C -d SalesKpiDW -Q "SELECT TOP 5 * FROM LoadLog ORDER BY LogID DESC; SELECT TOP 7 * FROM vw_DailyKpi ORDER BY [Date] DESC"
```

## 4. Power BI Desktop
1. **Obtenir les données → SQL Server** : serveur `localhost\SQLEXPRESS`, base `SalesKpiDW`, mode **Import**.
   Tables : `FactSales, DimCustomer, DimProduct, DimDate, FactTarget, LoadLog`.
2. Relations (plusieurs-à-un, sens unique) : `FactSales[DateKey]→DimDate[DateKey]`, `[CustomerKey]→DimCustomer`,
   `[ProductKey]→DimProduct`, `FactTarget[MonthStart]→DimDate[MonthStart]`. Marquer `DimDate` comme table de dates (`Date`).
3. Table vide `_Measures` (Entrer des données) puis coller les mesures de `powerbi\measures.dax` **une par une**
   (Revenue, Cost, Profit, Margin % en premier).
4. Construire les pages selon `powerbi\dashboard_layout.md` (la page 4 = tendance par jour ;
   filtre visuel `Show Date = 1` sur la courbe).

## 5. Actualisation automatique du dashboard
Power BI Desktop ne s'actualise pas seul. Publiez le rapport sur Power BI Service, puis :
- installez une **passerelle de données locale** (le SQL Server est sur votre PC) ;
- **Jeu de données → Paramètres → Actualisation planifiée** (par ex. 07:00, après le chargement de 06:30) ;
- ou renseignez les IDs dans `refresh_powerbi.ps1` pour déclencher l'actualisation juste après chaque chargement.

## 6. Dépannage
| Erreur | Solution |
|---|---|
| `Data source name not found` | nom du driver ODBC (17/18) dans `CONN` |
| `Invalid object name 'stg_...'` / `usp_LoadDay` | étape 2 (b, c) non exécutée |
| `Échec de INSERT ... options SET` | relancer `02_daily_load.sql` (il contient `SET QUOTED_IDENTIFIER ON`) |
| Date « skipped » alors qu'on veut la recharger | `DELETE FROM LoadLog WHERE LoadDate='AAAA-MM-JJ'` (les commandes déjà présentes ne sont pas dupliquées) |
| Courbe jusqu'au 31 décembre | filtre visuel `Show Date = 1` |

> Un fichier `.pbix` ne peut pas être généré par code : le rapport se construit dans Power BI Desktop
> à partir du modèle et des mesures fournis.
