# Importer le module ImportExcel
Import-Module ImportExcel

# Chemin vers le fichier Excel
$excelFilePath = "C:\chemin\vers\votre\fichier\s01_BillU (2).xlsx"

# Lire le fichier Excel
$excelData = Import-Excel -Path $excelFilePath

# URL de l'API REST de GLPI
$glpiApiUrl = "https://votre-instance-glpi/apirest.php"

# Clé API et identifiants
$apiKey = "votre-cle-api"
$appToken = "votre-app-token"
$sessionToken = ""

# Obtenir un token de session
$response = Invoke-RestMethod -Uri "$glpiApiUrl/initSession" -Method Post -Headers @{
    "Content-Type" = "application/json"
    "App-Token" = $appToken
    "Authorization" = "user_token $apiKey"
} -Body '{}'

$sessionToken = $response.session_token

# Parcourir les lignes du fichier Excel
foreach ($row in $excelData) {
    $departement = $row.'Département'
    $service = $row.'Service'
    $nomEntite = $row."Nom de l'entité"

    # Créer une entité via l'API REST
    $body = @{
        input = @{
            name = $nomEntite
            entities_id = 0 # ID de l'entité parente, si applicable
            comment = "$departement - $service"
        }
    }

    $response = Invoke-RestMethod -Uri "$glpiApiUrl/Entity" -Method Post -Headers @{
        "Content-Type" = "application/json"
        "Session-Token" = $sessionToken
        "App-Token" = $appToken
    } -Body ($body | ConvertTo-Json)

    Write-Output "Entité créée : $nomEntite"
}

# Fermer la session
Invoke-RestMethod -Uri "$glpiApiUrl/killSession" -Method Post -Headers @{
    "Content-Type" = "application/json"
    "Session-Token" = $sessionToken
    "App-Token" = $appToken
} -Body '{}'
