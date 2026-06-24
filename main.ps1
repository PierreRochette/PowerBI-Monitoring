# This section is dedicated to declaring functions
# Check a comment "MAIN" to see their orchestration


# GetPowerBiToken : 
# Get OAuth2 token for PBI API
# Prérequisite : l'App Registration has Power BI permissions configured in Entra 

function GetPowerBiToken {

    param(
        [string]$TenantId, # Tenant Azure AD 
        [string]$ClientId, # Identifiant public de l'application 
        [string]$ClientSecret # Secret de l'application
    )

    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        resource      = "https://analysis.windows.net/powerbi/api"
    }

    try {
        $response = Invoke-RestMethod `
            -Method POST `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/token" `
            -Body $body

        

        return $response.access_token
    
    }
    catch {
  
        Write-Error "Échec obtention du token : $_"
        exit
    }

}

# GetPowerBIActivityEvents : 
# All ActivityEvents in the admin endpoint
# Check Microsoft offical doc for further infos
# https://learn.microsoft.com/fr-fr/rest/api/power-bi/admin/get-activity-events

function GetPowerBIActivityEvents {

    param(
        [string]$Token, 
        [string]$StartDate, 
        [string]$EndDate, # yyyy-MM-ddTHH:59:59.000Z
        [string]$Filter = "" # Optionnal
    )

    $uri = "https://api.powerbi.com/v1.0/myorg/admin/activityevents?startDateTime='$StartDate'&endDateTime='$EndDate'"

    if($Filter -ne "") {
        $uri += "&`$filter=$Filter"
    }

    $allEvents = @()

    # Pagination loop

    do {

        $response = Invoke-RestMethod `
            -Uri $uri `
            -Headers @{ Authorization = "Bearer $Token"}

        $allEvents += $response.activityEventEntities

        Write-Output "Page récupérée : $($response.activityEventEntities.Count)"

        $uri = $response.continuationUri

    } while ($response.lastResultSet -eq $false)

    Write-Output "Total events : $($allEvents.Count)"
    return $allEvents

}

function Invoke-SqlQuery {

    param(
        [string]$Server, # Serveur Azure SQL 
        [string]$Database, # DB cible
        [string]$Token, # Token SQL, voire fonction Get-SqlToken
        [string]$Query # Requête SQL à exécuter
    )

    $conn = New-Object System.Data.SqlClient.SqlConnection

    $conn.ConnectionString = "Server=$Server;Database=$Database;Encrypt=True;TrustServerCertificate=False;"
    
    # Pour avoir un token, voir la fonction Get-SqlToken
    $conn.AccessToken = $Token
    $conn.Open()

    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $Query
    $cmd.CommandTimeout = 120
    $cmd.ExecuteNonQuery() | Out-Null

    $conn.Close()

}

# Token OAuth2 for Azure SQL 

function Get-SqlToken {

    param(
        [string]$TenantId, 
        [string]$ClientId, 
        [string]$ClientSecret
    )

    $body = @{
        grant_type      = "client_credentials"
        client_id       = $ClientId
        client_secret   = $ClientSecret
        scope           = "https://database.windows.net/.default"
    }

    try {

        $response = Invoke-RestMethod `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -Method Post `
            -Body $body

        return $response.access_token

    } catch {

        Write-Error "Échec obtention du token SQL : $_"
        exit

    }

}

function Initialize-Database {

    param(
        [string]$Server,
        [string]$Database, 
        [string]$Token
    )

    $query = @"
    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ActivityEvents')
BEGIN
    CREATE TABLE ActivityEvents (
        Id                       INT            NOT NULL IDENTITY(1,1) CONSTRAINT PK_ActivityEvents PRIMARY KEY, 
        EventId                       NVARCHAR(100)  NULL,
        RecordType               INT            NULL,
        CreationTime             DATETIME2      NULL,
        Operation                NVARCHAR(150)  NULL,
        Activity                 NVARCHAR(150)  NULL,
        OrganizationId           NVARCHAR(100)  NULL,
        UserId                   NVARCHAR(255)  NULL,
        UserKey                  NVARCHAR(100)  NULL,
        UserType                 INT            NULL,
        Workload                 NVARCHAR(100)  NULL,
        ResultStatus             NVARCHAR(50)   NULL,
        ClientIP                 NVARCHAR(50)   NULL,
        UserAgent                NVARCHAR(500)  NULL,
        WorkspaceId              NVARCHAR(100)  NULL,
        WorkSpaceName            NVARCHAR(255)  NULL,
        ObjectId                 NVARCHAR(100)  NULL,
        ObjectType               NVARCHAR(100)  NULL,
        ObjectDisplayName        NVARCHAR(255)  NULL,
        RequestId                NVARCHAR(100)  NULL,
        Experience               NVARCHAR(100)  NULL,
        DatasetId                NVARCHAR(100)  NULL,
        DatasetName              NVARCHAR(255)  NULL,
        ItemName                 NVARCHAR(255)  NULL,
        ItemId                   NVARCHAR(100)  NULL,
        ArtifactId               NVARCHAR(100)  NULL,
        ArtifactName             NVARCHAR(255)  NULL,
        ArtifactKind             NVARCHAR(100)  NULL,
        DataConnectivityMode     NVARCHAR(100)  NULL,
        RefreshType              NVARCHAR(100)  NULL,
        IsSuccess                BIT            NULL,
        LastRefreshTime          DATETIME2      NULL,
        ActivityId               NVARCHAR(100)  NULL,
        CapacityId               NVARCHAR(100)  NULL,
        CapacityName             NVARCHAR(255)  NULL,
        RefreshEnforcementPolicy INT            NULL,
        BillingType              INT            NULL,
        RawJson                  NVARCHAR(MAX)  NULL,
        InsertedAt               DATETIME2      NOT NULL DEFAULT GETUTCDATE()
    )
END
"@

    Invoke-SqlQuery -Server $Server -Database $Database -Token $Token -Query $query

}

function WriteEventsToDatabase {

    param(
        [array]$Events, #Tableau d'évènements issu de GetPowerBiActivityEvents
        [string]$Server, 
        [string]$Database, 
        [string]$Token
    )

    $dt = New-Object System.Data.DataTable

    $cols = @(
        @{ Name = "EventId";                       Type = [string] }
        @{ Name = "RecordType";               Type = [object] }
        @{ Name = "CreationTime";             Type = [object] }
        @{ Name = "Operation";                Type = [string] }
        @{ Name = "Activity";                 Type = [string] }
        @{ Name = "OrganizationId";           Type = [string] }
        @{ Name = "UserId";                   Type = [string] }
        @{ Name = "UserKey";                  Type = [string] }
        @{ Name = "UserType";                 Type = [object] }
        @{ Name = "Workload";                 Type = [string] }
        @{ Name = "ResultStatus";             Type = [string] }
        @{ Name = "ClientIP";                 Type = [string] }
        @{ Name = "UserAgent";                Type = [string] }
        @{ Name = "WorkspaceId";              Type = [string] }
        @{ Name = "WorkSpaceName";            Type = [string] }
        @{ Name = "ObjectId";                 Type = [string] }
        @{ Name = "ObjectType";               Type = [string] }
        @{ Name = "ObjectDisplayName";        Type = [string] }
        @{ Name = "RequestId";                Type = [string] }
        @{ Name = "Experience";               Type = [string] }
        @{ Name = "DatasetId";                Type = [string] }
        @{ Name = "DatasetName";              Type = [string] }
        @{ Name = "ItemName";                 Type = [string] }
        @{ Name = "ItemId";                   Type = [string] }
        @{ Name = "ArtifactId";               Type = [string] }
        @{ Name = "ArtifactName";             Type = [string] }
        @{ Name = "ArtifactKind";             Type = [string] }
        @{ Name = "DataConnectivityMode";     Type = [string] }
        @{ Name = "RefreshType";              Type = [string] }
        @{ Name = "IsSuccess";                Type = [object] }
        @{ Name = "LastRefreshTime";          Type = [object] }
        @{ Name = "ActivityId";               Type = [string] }
        @{ Name = "CapacityId";               Type = [string] }
        @{ Name = "CapacityName";             Type = [string] }
        @{ Name = "RefreshEnforcementPolicy"; Type = [object] }
        @{ Name = "BillingType";              Type = [object] }
        @{ Name = "RawJson";                  Type = [string] }
    )

    foreach ($col in $cols) {
        $dt.Columns.Add($col.Name, $col.Type) | Out-Null
    }

    foreach ($event in $Events) {

        $row = $dt.NewRow()


        $row["EventId"]                  = if ($event.Id)                                  { $event.Id }                       else { [DBNull]::Value }
        $row["RecordType"]               = if ($null -ne $event.RecordType)               { $event.RecordType }               else { [DBNull]::Value }
        $row["CreationTime"]             = if ($event.CreationTime)                        { [datetime]$event.CreationTime }   else { [DBNull]::Value }
        $row["Operation"]                = if ($event.Operation)                           { $event.Operation }                else { [DBNull]::Value }
        $row["Activity"]                 = if ($event.Activity)                            { $event.Activity }                 else { [DBNull]::Value }
        $row["OrganizationId"]           = if ($event.OrganizationId)                      { $event.OrganizationId }           else { [DBNull]::Value }
        $row["UserId"]                   = if ($event.UserId)                              { $event.UserId }                   else { [DBNull]::Value }
        $row["UserKey"]                  = if ($event.UserKey)                             { $event.UserKey }                  else { [DBNull]::Value }
        $row["UserType"]                 = if ($null -ne $event.UserType)                  { $event.UserType }                 else { [DBNull]::Value }
        $row["Workload"]                 = if ($event.Workload)                            { $event.Workload }                 else { [DBNull]::Value }
        $row["ResultStatus"]             = if ($event.ResultStatus)                        { $event.ResultStatus }             else { [DBNull]::Value }
        $row["ClientIP"]                 = if ($event.ClientIP)                            { $event.ClientIP }                 else { [DBNull]::Value }
        $row["UserAgent"]                = if ($event.UserAgent)                           { $event.UserAgent }                else { [DBNull]::Value }
        $row["WorkspaceId"]              = if ($event.WorkspaceId)                         { $event.WorkspaceId }              else { [DBNull]::Value }
        $row["WorkSpaceName"]            = if ($event.WorkSpaceName)                       { $event.WorkSpaceName }            else { [DBNull]::Value }
        $row["ObjectId"]                 = if ($event.ObjectId)                            { $event.ObjectId }                 else { [DBNull]::Value }
        $row["ObjectType"]               = if ($event.ObjectType)                          { $event.ObjectType }               else { [DBNull]::Value }
        $row["ObjectDisplayName"]        = if ($event.ObjectDisplayName)                   { $event.ObjectDisplayName }        else { [DBNull]::Value }
        $row["RequestId"]                = if ($event.RequestId)                           { $event.RequestId }                else { [DBNull]::Value }
        $row["Experience"]               = if ($event.Experience)                          { $event.Experience }               else { [DBNull]::Value }
        $row["DatasetId"]                = if ($event.DatasetId)                           { $event.DatasetId }                else { [DBNull]::Value }
        $row["DatasetName"]              = if ($event.DatasetName)                         { $event.DatasetName }              else { [DBNull]::Value }
        $row["ItemName"]                 = if ($event.ItemName)                            { $event.ItemName }                 else { [DBNull]::Value }
        $row["ItemId"]                   = if ($event.ItemId)                              { $event.ItemId }                   else { [DBNull]::Value }
        $row["ArtifactId"]               = if ($event.ArtifactId)                         { $event.ArtifactId }               else { [DBNull]::Value }
        $row["ArtifactName"]             = if ($event.ArtifactName)                        { $event.ArtifactName }             else { [DBNull]::Value }
        $row["ArtifactKind"]             = if ($event.ArtifactKind)                        { $event.ArtifactKind }             else { [DBNull]::Value }
        $row["DataConnectivityMode"]     = if ($event.DataConnectivityMode)                { $event.DataConnectivityMode }     else { [DBNull]::Value }
        $row["RefreshType"]              = if ($event.RefreshType)                         { $event.RefreshType }              else { [DBNull]::Value }
        $row["IsSuccess"]                = if ($null -ne $event.IsSuccess)                 { [bool]$event.IsSuccess }          else { [DBNull]::Value }
        $row["LastRefreshTime"]          = if ($event.LastRefreshTime)                     { [datetime]$event.LastRefreshTime } else { [DBNull]::Value }
        $row["ActivityId"]               = if ($event.ActivityId)                          { $event.ActivityId }               else { [DBNull]::Value }
        $row["CapacityId"]               = if ($event.CapacityId)                          { $event.CapacityId }               else { [DBNull]::Value }
        $row["CapacityName"]             = if ($event.CapacityName)                        { $event.CapacityName }             else { [DBNull]::Value }
        $row["RefreshEnforcementPolicy"] = if ($null -ne $event.RefreshEnforcementPolicy)  { $event.RefreshEnforcementPolicy } else { [DBNull]::Value }
        $row["BillingType"]              = if ($null -ne $event.BillingType)               { $event.BillingType }              else { [DBNull]::Value }
        $row["RawJson"]                  = $event | ConvertTo-Json -Compress -Depth 5

        $dt.Rows.Add($row)

    }

    Write-Host "Data to insert : $($dt.Rows.Count) lines"

    $conn = New-Object System.Data.SqlClient.SqlConnection
    $conn.ConnectionString = "Server=$Server;Database=$Database;Encrypt=True;TrustServerCertificate=False;"
    $conn.AccessToken = $Token
    $conn.Open()

    $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy($conn)
    $bulkCopy.DestinationTableName = "ActivityEvents"
    $bulkCopy.BulkCopyTimeout = 300

    $bulkCopy.ColumnMappings.Add("EventId",                  "EventId")                  | Out-Null
    $bulkCopy.ColumnMappings.Add("RecordType",               "RecordType")               | Out-Null
    $bulkCopy.ColumnMappings.Add("CreationTime",             "CreationTime")             | Out-Null
    $bulkCopy.ColumnMappings.Add("Operation",                "Operation")                | Out-Null
    $bulkCopy.ColumnMappings.Add("Activity",                 "Activity")                 | Out-Null
    $bulkCopy.ColumnMappings.Add("OrganizationId",           "OrganizationId")           | Out-Null
    $bulkCopy.ColumnMappings.Add("UserId",                   "UserId")                   | Out-Null
    $bulkCopy.ColumnMappings.Add("UserKey",                  "UserKey")                  | Out-Null
    $bulkCopy.ColumnMappings.Add("UserType",                 "UserType")                 | Out-Null
    $bulkCopy.ColumnMappings.Add("Workload",                 "Workload")                 | Out-Null
    $bulkCopy.ColumnMappings.Add("ResultStatus",             "ResultStatus")             | Out-Null
    $bulkCopy.ColumnMappings.Add("ClientIP",                 "ClientIP")                 | Out-Null
    $bulkCopy.ColumnMappings.Add("UserAgent",                "UserAgent")                | Out-Null
    $bulkCopy.ColumnMappings.Add("WorkspaceId",              "WorkspaceId")              | Out-Null
    $bulkCopy.ColumnMappings.Add("WorkSpaceName",            "WorkSpaceName")            | Out-Null
    $bulkCopy.ColumnMappings.Add("ObjectId",                 "ObjectId")                 | Out-Null
    $bulkCopy.ColumnMappings.Add("ObjectType",               "ObjectType")               | Out-Null
    $bulkCopy.ColumnMappings.Add("ObjectDisplayName",        "ObjectDisplayName")        | Out-Null
    $bulkCopy.ColumnMappings.Add("RequestId",                "RequestId")                | Out-Null
    $bulkCopy.ColumnMappings.Add("Experience",               "Experience")               | Out-Null
    $bulkCopy.ColumnMappings.Add("DatasetId",                "DatasetId")                | Out-Null
    $bulkCopy.ColumnMappings.Add("DatasetName",              "DatasetName")              | Out-Null
    $bulkCopy.ColumnMappings.Add("ItemName",                 "ItemName")                 | Out-Null
    $bulkCopy.ColumnMappings.Add("ItemId",                   "ItemId")                   | Out-Null
    $bulkCopy.ColumnMappings.Add("ArtifactId",               "ArtifactId")               | Out-Null
    $bulkCopy.ColumnMappings.Add("ArtifactName",             "ArtifactName")             | Out-Null
    $bulkCopy.ColumnMappings.Add("ArtifactKind",             "ArtifactKind")             | Out-Null
    $bulkCopy.ColumnMappings.Add("DataConnectivityMode",     "DataConnectivityMode")     | Out-Null
    $bulkCopy.ColumnMappings.Add("RefreshType",              "RefreshType")              | Out-Null
    $bulkCopy.ColumnMappings.Add("IsSuccess",                "IsSuccess")                | Out-Null
    $bulkCopy.ColumnMappings.Add("LastRefreshTime",          "LastRefreshTime")          | Out-Null
    $bulkCopy.ColumnMappings.Add("ActivityId",               "ActivityId")               | Out-Null
    $bulkCopy.ColumnMappings.Add("CapacityId",               "CapacityId")               | Out-Null
    $bulkCopy.ColumnMappings.Add("CapacityName",             "CapacityName")             | Out-Null
    $bulkCopy.ColumnMappings.Add("RefreshEnforcementPolicy", "RefreshEnforcementPolicy") | Out-Null
    $bulkCopy.ColumnMappings.Add("BillingType",              "BillingType")              | Out-Null
    $bulkCopy.ColumnMappings.Add("RawJson",                  "RawJson")                  | Out-Null


    $bulkCopy.WriteToServer($dt)

    $conn.Close()

    Write-Host "Insert finished : $($dt.Rows.Count) events insérés."

}

function Remove-OldRecords {

    param(
        [string]$Server,
        [string]$Database, 
        [string]$Token, 
        [int]$RetentionDays
    )

    $cutoffDate = (Get-Date).ToUniversalTime().AddDays(-$RetentionDays).ToString("yyyy-MM-dd")

    $query = @"
DELETE FROM ActivityEvents
WHERE InsertedAt < '$cutoffDate'
"@

    Invoke-SqlQuery -Server $Server -Database $Database -Token $Token -Query $query

}

# MAIN 

# 1. Getting Automation Account variables

$tenantId       = Get-AutomationVariable -Name "PBI-TenantId"
$clientId       = Get-AutomationVariable -Name "PBI-ClientId"
$clientSecret   = Get-AutomationVariable -Name "PBI-ClientSecret"
$sqlServer      = Get-AutomationVariable -Name "SQL-Server"
$sqlDatabase    = Get-AutomationVariable -Name "SQL-Database"

# 2. Get tokens (PBI & SQL Server)

$pbiToken = GetPowerBiToken -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret

$sqlToken = Get-SqlToken -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret


# 3. Fetching datas from day -1 

$startDate = (Get-Date).AddDays(-1).ToString("yyyy-MM-ddT00:00:00.000Z")
$endDate   = (Get-Date).AddDays(-1).ToString("yyyy-MM-ddT23:59:59.000Z")

# For debugging 

# $startDate = "2026-06-19T00:00:00.000Z"
# $endDate   = "2026-06-19T23:59:59.000Z"

# 4. PBI API call 
# Saving the response in events variable

$events = GetPowerBIActivityEvents `
    -Token $pbiToken `
    -StartDate $startDate `
    -EndDate $endDate `
    -Filter ""

# 4 bis - Stopping script if no valid SQL token

if ($sqlToken.Length -eq 0) {
    Write-Error "SQL Token vide - arrêt du runbook"
    exit
}

# 5. Initialize database 

Initialize-Database -Server $sqlServer -Database $sqlDatabase -Token $sqlToken

# 6. Retention

Remove-OldRecords -Server $sqlServer -Database $sqlDatabase -Token $sqlToken -RetentionDays 90

# 7. Insert events 

WriteEventsToDatabase -Events $events -Server $sqlServer -Database $sqlDatabase -Token $sqlToken