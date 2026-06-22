# DECLARATION DES FONCTIONS 
# LES FONCTIONS POWER BI ET SQL SONT TOUTES DEFINIES
# POUR L'ORCHESTRATION DE LEUR EXECUTION, CHERCHER LA SECTION COMMENTEE 'MAIN'
# EN CAPITALE COMME CETTE EN-TÊTE

## FONCTIONS POWER BI 

# GetPowerBiToken : 
# Récupère un token OAuth2 pour l'API Power BI via le flux Client Credentials.
# Ce flux est conçu pour les applications s'authentifiant sans utilisateur 
# (service-to-service).
# Prérequis : l'App Registration doit avoir les permissions API Power BI configurées 
# dans Entra.

function GetPowerBiToken {

    # Paramètres de la fonction à appeler

    param(
        [string]$TenantId, # Tenant Azure AD 
        [string]$ClientId, # Identifiant public de l'application 
        [string]$ClientSecret # Secret de l'application
    )

    # Payload pour la requête API

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

        # On retourne uniquement le access Token après avoir appelé le endpoint

        return $response.access_token
    
    }
    catch {
        # Log en cas d'échec
        Write-Error "Échec obtention du token : $_"
        exit
    }

}

# GetPowerBIActivityEvents : 
# Récupère tous les événements d'activité Power BI sur une plage de dates donnée.
# L'API pagine les résultats — cette fonction gère automatiquement la pagination
# via continuationUri jusqu'à ce que lastResultSet soit true.
# Contrainte API : la plage StartDate/EndDate ne peut pas dépasser 24 heures.

function GetPowerBIActivityEvents {

    param(
        [string]$Token, # Token PowerBI issu de l'appel à la fonction GetPowerBiToken
        [string]$StartDate, # Date de début, gérée par une fonction SQL dans le main
        [string]$EndDate, # yyyy-MM-ddTHH:59:59.000Z
        [string]$Filter = "" # Optionnel pour filtrer sur le type d'activité Filtre OData optionnel, ex: "Activity eq 'ViewReport'"
    )

    # Les dates doivent être entre guillements simples dans la requête
    $uri = "https://api.powerbi.com/v1.0/myorg/admin/activityevents?startDateTime='$StartDate'&endDateTime='$EndDate'"

    if($Filter -ne "") {
        $uri += "&`$filter=$Filter"
    }

    $allEvents = @()

    # Boucle de pagination : l'API retourne jusqu'à 5000 events par page.
    # À chaque itération, continuationUri pointe vers la page suivante.
    # La boucle s'arrête quand lastResultSet passe à true.

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

## FONCTIONS SQL

# Invoke-SqlQuery
# Exécute une requête SQL sans retour de données (DDL ou DML : CREATE, INSERT, 
# UPDATE, DELETE).
# Pour des requêtes SELECT, utiliser ExecuteReader à la place de ExecuteNonQuery.
# Cette fonction sera ensuite utilisée pour initialiser la database et inscrire les données

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

# Get-SqlToken
# Récupère un token OAuth2 pour Azure SQL via le flux Client Credentials.

function Get-SqlToken {

    param(
        [string]$TenantId, 
        [string]$ClientId, 
        [string]$ClientSecret
    )

    # Le scope "/.default" demande 
    # toutes les permissions déjà accordées à l'App Registration.
    # L'App Registration doit être membre d'un rôle sur la base SQL 
    # (db_datareader, db_datawriter...).

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

#Initialize-Database
# Crée la table Bronze ActivityEvents si elle n'existe pas encore.
# Idempotente : peut être appelée à chaque exécution du runbook sans risque.

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

# WriteEventsToDatabase
# Insère tous les événements en base via SqlBulkCopy (insertion en batch).
# Plus performant que des INSERT un par un — recommandé pour des volumes > 100 lignes.

function WriteEventsToDatabase {

    param(
        [array]$Events, #Tableau d'évènements issu de GetPowerBiActivityEvents
        [string]$Server, 
        [string]$Database, 
        [string]$Token
    )

    # Construction du DataTable en mémoire avant l'insertion bulk.
    # Les colonnes de type [object] acceptent à la fois des valeurs typées et DBNull,
    # ce qui est nécessaire pour les colonnes INT, BIT et DATETIME2 nullable en SQL.
    # Les colonnes de type [string] sont pour les NVARCHAR.

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

        # Règle de mapping null → DBNull :
        # - Chaînes (NVARCHAR)  : if ($event.Champ) — null et chaîne vide sont falsy
        # - Numériques/booléens : if ($null -ne $event.Champ) — évite de traiter 0 ou $false comme null

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

    # Mapping explicite colonne DataTable → colonne SQL.
    # Obligatoire ici car la table SQL contient "Id" (IDENTITY) en première position :
    # sans mapping, SqlBulkCopy mappe par position et tente d'écrire dans Id, ce qui échoue.
    # "Id" et "InsertedAt" sont exclus — gérés automatiquement par SQL Server.

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

# SUPPRESSION DES ANCIENNES DONNEES
# Déclaration de fonction

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
# ORCHESTRATION DES DIFFERENTES FONCTIONS CREES 

# 1. Récupération des variables du compte d'automation

$tenantId       = Get-AutomationVariable -Name "PBI-TenantId"
$clientId       = Get-AutomationVariable -Name "PBI-ClientId"
$clientSecret   = Get-AutomationVariable -Name "PBI-ClientSecret"
$sqlServer      = Get-AutomationVariable -Name "SQL-Server"
$sqlDatabase    = Get-AutomationVariable -Name "SQL-Database"

Write-Host "Variables OK"
Write-Host "Getting tokens...."

# 2. Récupérer le token power bi et le token sql server

$pbiToken = GetPowerBiToken -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret
Write-Host "Power BI Token OK"

$sqlToken = Get-SqlToken -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret
Write-Host "SQL Token OK"

Write-Host "Calling API..."

# 3. Générer start date et end date (J-1 entre 00:00 et 23:59:59)

$startDate = (Get-Date).AddDays(-1).ToString("yyyy-MM-ddT00:00:00.000Z")
$endDate   = (Get-Date).AddDays(-1).ToString("yyyy-MM-ddT23:59:59.000Z")

# Pour debugger
# Ne pas utiliser de valeurs en dur sur les dates pour la prod

# $startDate = "2026-06-19T00:00:00.000Z"
# $endDate   = "2026-06-19T23:59:59.000Z"

# 4. Call API PowerBI avec token et dates, sans filtre 
# Enregistrement de ce qui est récupéré dans une variable events

$events = GetPowerBIActivityEvents `
    -Token $pbiToken `
    -StartDate $startDate `
    -EndDate $endDate `
    -Filter ""

Write-Host "Total Events : $($events.Count)"

Write-Host "SQL Token length: $($sqlToken.Length)"

# 4 bis - Arrêt du script si pas de token SQL Server valid

if ($sqlToken.Length -eq 0) {
    Write-Error "SQL Token vide - arrêt du runbook"
    exit
}

Write-Host "Checking database..."

# 5. Initialisation de la base de données si nécéssaire 

Initialize-Database -Server $sqlServer -Database $sqlDatabase -Token $sqlToken

Write-Host "Insertion en base..."

# 6. Nettoyage des données

Remove-OldRecords -Server $sqlServer -Database $sqlDatabase -Token $sqlToken -RetentionDays 3 

# 7. Insertion des évènements en appelant la fonction WriteEventsToDatabase

WriteEventsToDatabase -Events $events -Server $sqlServer -Database $sqlDatabase -Token $sqlToken

Write-Output "Runbook done."