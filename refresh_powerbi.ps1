# Triggers a refresh of the published dataset in Power BI Service (REST API).
# Needs: Power BI Pro/PPU, dataset published, data gateway online for the local SQL Server.
# Fill in the two IDs (Power BI Service > dataset > Settings: they are in the URL).
#   .../groups/<WORKSPACE_ID>/datasets/<DATASET_ID>/...
# First time only:  Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser ; Connect-PowerBIServiceAccount
$WorkspaceId = "PUT-WORKSPACE-ID-HERE"
$DatasetId   = "PUT-DATASET-ID-HERE"

if ($WorkspaceId -like "PUT-*") { Write-Host "refresh_powerbi.ps1 not configured -> skipped (use scheduled refresh in the service instead)"; return }

Import-Module MicrosoftPowerBIMgmt.Profile
Connect-PowerBIServiceAccount | Out-Null     # unattended use: add -ServicePrincipal -Credential -Tenant
Invoke-PowerBIRestMethod -Method Post -Url "groups/$WorkspaceId/datasets/$DatasetId/refreshes" -Body '{"notifyOption":"NoNotification"}'
Write-Host "Power BI refresh requested."
