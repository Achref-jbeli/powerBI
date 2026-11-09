# Power BI – Excel → SQL Server → KPI → Dashboard (chargement quotidien automatique)

Flux complet **sans Python ni pip** : PowerShell + ADO.NET natif + SQL Server Agent.

```
inbox\2026-11-08\sales_2026-11-08.xlsx      (un dossier par jour, nom = AAAA-MM-JJ)
        │  SQL Server Agent Job (06:30) → run_daily.ps1 → load_daily.ps1
        │    • date déjà chargée (LoadLog = Success) → on ne fait rien
        │    • date nouvelle → stg_* → usp_LoadDay (incrémental) → LoadLog
        ▼
SQL Server « SalesKpiDW » : schéma en étoile + vues vw_MonthlyKpi, vw_DailyKpi, vw_CustomerRfm
        │  refresh_powerbi.ps1 (uniquement si nouvelles données)
        ▼
Power BI : modèle + mesures DAX (powerbi/measures.dax) + pages (powerbi/dashboard_layout.md)
```

## Prérequis
| Outil | Remarque |
|---|---|
| SQL Server (Express suffit) + SSMS ou `sqlcmd` | serveur par défaut : `localhost\SQLEXPRESS` |
| PowerShell 5.1+ | déjà présent sur Windows 10/11 |
| Microsoft ACE OleDb 2016 | lit les fichiers Excel sans Python — [télécharger ici](https://www.microsoft.com/en-us/download/details.aspx?id=54920) (gratuit, sans pip) |
| Microsoft Excel | uniquement pour `generate_sample_excel.ps1` (données de test) |
| Power BI Desktop | gratuit (Microsoft Store) |

> **SQL Server Express** n'inclut pas SQL Server Agent.
> Utilisez `register_daily_task.ps1` (Planificateur de tâches Windows) à la place de `sql\03_agent_job.sql`.

---

## 1. Installation (une seule fois)

```powershell
cd C:\Users\achre\Downloads\PowerBI-Sales-Example

# Tout en une commande :
.\setup.ps1

# Ou étape par étape :
sqlcmd -S "localhost\SQLEXPRESS" -E -C -i sql\01_schema.sql   # crée la base et les tables
sqlcmd -S "localhost\SQLEXPRESS" -E -C -i sql\02_daily_load.sql  # LoadLog + usp_LoadDay + vw_DailyKpi
.\generate_sample_excel.ps1    # crée data\sales_data.xlsx (données de test, nécessite Excel)
.\load_excel_to_db.ps1         # charge l'historique initial
```

Un autre serveur ? Modifiez `$Server` en haut de chaque script `.ps1` ou passez `-Server "..."`.

---

## 2. Chargement quotidien

### Structure des dossiers
```
inbox\
  2026-11-08\
    sales_2026-11-08.xlsx      ← feuille "Sales" obligatoire
  2026-11-09\
    sales_2026-11-09.xlsx
```

Feuille **Sales** requise : `OrderID, OrderDate, CustomerID, ProductID, Quantity, DiscountPct`
Feuilles facultatives : `Customers`, `Products`, `Targets`

### Lancer manuellement
```powershell
.\load_daily.ps1
# [2026-11-08] chargé : 95 nouvelles commandes depuis sales_2026-11-08.xlsx
# [2026-11-07] déjà chargé -> ignoré
```
Codes de sortie : `0` nouvelles données · `2` rien de nouveau · `1` échec (voir `logs\` et `LoadLog`)

### Automatiser avec SQL Server Agent (Standard/Developer)
```powershell
# Crée le job "SalesKpi_DailyLoad" à 06:30
sqlcmd -S "localhost\SQLEXPRESS" -E -C -i sql\03_agent_job.sql
```

### Automatiser avec le Planificateur de tâches (Express)
```powershell
.\register_daily_task.ps1 -Time 06:30
```

### Vérifier
```powershell
sqlcmd -S "localhost\SQLEXPRESS" -E -C -d SalesKpiDW -Q "SELECT TOP 5 * FROM LoadLog ORDER BY LogID DESC; SELECT TOP 7 * FROM vw_DailyKpi ORDER BY [Date] DESC"
```

---

## 3. Power BI Desktop

1. **Obtenir les données → SQL Server** : serveur `localhost\SQLEXPRESS`, base `SalesKpiDW`, mode **Import**.
   Tables : `FactSales, DimCustomer, DimProduct, DimDate, FactTarget, LoadLog`.
2. **Relations** (plusieurs-à-un, sens unique) :
   - `FactSales[DateKey]` → `DimDate[DateKey]`
   - `FactSales[CustomerKey]` → `DimCustomer[CustomerKey]`
   - `FactSales[ProductKey]` → `DimProduct[ProductKey]`
   - `FactTarget[MonthStart]` → `DimDate[MonthStart]`
   - Marquer `DimDate` comme table de dates (colonne `Date`).
3. **Mesures** : créer une table vide `_Measures`, puis coller les mesures de `powerbi\measures.dax` une par une
   (commencer par `Revenue`, `Cost`, `Profit`, `Margin %`).
4. **Pages** : suivre `powerbi\dashboard_layout.md` (page 4 = tendance quotidienne ;
   filtre visuel `Show Date = 1` sur la courbe).

---

## 4. Actualisation automatique du dashboard

| Scénario | Solution |
|---|---|
| Power BI Desktop seulement | Cliquer **Actualiser** après chaque chargement |
| Power BI Service (automatique) | Publier + passerelle de données locale + actualisation planifiée à 07:00 |
| Déclenchement immédiat après chargement | Renseigner les IDs dans `refresh_powerbi.ps1` |

---

## 5. Dépannage

| Erreur | Solution |
|---|---|
| `The 'Microsoft.ACE.OLEDB.12.0' provider is not registered` | Installer [ACE OleDb 2016](https://www.microsoft.com/en-us/download/details.aspx?id=54920) |
| `Invalid object name 'stg_...'` | Relancer `sql\01_schema.sql` dans SSMS |
| `Échec de INSERT ... options SET` | Relancer `sql\02_daily_load.sql` |
| Date « ignorée » à recharger | `DELETE FROM LoadLog WHERE LoadDate='AAAA-MM-JJ'` dans SSMS |
| SQL Server Agent absent | Utiliser `register_daily_task.ps1` (Planificateur Windows) |
| Courbe Power BI jusqu'au 31 décembre | Ajouter filtre visuel `Show Date = 1` |

> Un fichier `.pbix` ne peut pas être généré par script : le rapport se construit dans Power BI Desktop
> à partir du modèle et des mesures fournis dans `powerbi\`.
