function New-EntraIDTAP {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserId,                                  # Identifiant de l'utilisateur (obligatoire)
        [int]$DurationInMinutes = 180,                    # Durée en minutes (3 heures par défaut)
        [bool]$IsUsableOnce = $false                      # Indique si le TAP est à usage unique
    )

    # Corps de la requête pour le TAP.
    $Body = @{
        lifetimeInMinutes = $DurationInMinutes
        isUsableOnce = $IsUsableOnce
    } | ConvertTo-Json

    # URL de la requête TAP.
    $Uri = "https://graph.microsoft.com/v1.0/users/$UserId/authentication/temporaryAccessPassMethods"

    # Envoi de la requête TAP.
    try {
        $Result = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $Body -ErrorAction Stop

        # Retour du résultat.
        return $Result.temporaryAccessPass
    } catch {
        Write-Host -ForegroundColor Red "Erreur lors de la génération du TAP pour $UserId : $_"
        return $null
    }
}

# Connect to Graph (assure connection avant le traitement en lot)
if (-not (Get-MgContext)) {
    Connect-MgGraph -NoWelcome -Scopes "User.Read.All", "UserAuthenticationMethod.ReadWrite.All"
}

$users = (Import-csv -Path "users.csv").UserName
$hash = @{} 

# Traitement en boucle pour chaque utilisateur
ForEach ($user in $users) {
    $tap = New-EntraIDTAP -UserId $user
    if ($tap) {
        $hash.add($user, $tap)
    } else {
        Write-Host -ForegroundColor Yellow "Le TAP n'a pas pu être généré pour l'utilisateur $user."
    }
}

# Enregistrement des résultats dans un fichier CSV
$outpath = "Results.csv"
$hash.GetEnumerator() | Select-Object -Property @{N='User Name';E={$_.Key}}, @{N='Temporary Access Pass';E={$_.Value}} | Export-csv -Path $outpath -NoTypeInformation

$fullOutPath = (Resolve-Path -Path $outpath).Path
Write-Host "Les résultats ont été enregistrés dans:"
Write-Host -ForegroundColor Green $fullOutPath

