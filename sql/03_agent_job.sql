/* Step 3 - Crée un job SQL Server Agent qui lance run_daily.ps1 chaque jour à 06:30.
   Prérequis : SQL Server Agent doit être démarré (service SQLServerAgent$SQLEXPRESS).
   Adaptez @project_path si votre dossier est différent.                              */
USE msdb;
GO

DECLARE @project_path NVARCHAR(260) = N'C:\Users\achre\Downloads\PowerBI-Sales-Example'

-- Supprime le job s'il existe déjà
IF EXISTS (SELECT 1 FROM msdb.dbo.sysjobs WHERE name = N'SalesKpi_DailyLoad')
    EXEC sp_delete_job @job_name = N'SalesKpi_DailyLoad', @delete_unused_schedule = 1;

-- Création du job
EXEC sp_add_job
    @job_name       = N'SalesKpi_DailyLoad',
    @description    = N'Charge les nouveaux fichiers Excel quotidiens dans SalesKpiDW',
    @enabled        = 1;

-- Étape 1 : exécuter le script PowerShell
EXEC sp_add_jobstep
    @job_name        = N'SalesKpi_DailyLoad',
    @step_name       = N'Charger Excel',
    @subsystem       = N'PowerShell',
    @command         = N'Set-ExecutionPolicy -Scope Process Bypass; & "C:\Users\achre\Downloads\PowerBI-Sales-Example\run_daily.ps1"',
    @on_success_action = 1,  -- Quitter avec succès
    @on_fail_action    = 2;  -- Quitter avec échec

-- Planification : tous les jours à 06:30
EXEC sp_add_schedule
    @schedule_name  = N'Quotidien_0630',
    @freq_type      = 4,     -- Quotidien
    @freq_interval  = 1,
    @active_start_time = 063000;  -- 06:30:00

EXEC sp_attach_schedule @job_name = N'SalesKpi_DailyLoad', @schedule_name = N'Quotidien_0630';

-- Associe au serveur local
EXEC sp_add_jobserver @job_name = N'SalesKpi_DailyLoad', @server_name = N'(LOCAL)';

SELECT 'Job SalesKpi_DailyLoad créé, planifié à 06:30 chaque jour.' AS Résultat;
GO
