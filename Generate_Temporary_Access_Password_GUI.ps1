# Charger l'assembly de Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Chemin du logo (image)
$logoPath = "C:\Users\Dimby\Pictures\logo.jpg"  #Remplacez par le chemin de votre logo

# Chemins des fichiers de paramètres et de log
$settingsFilePath = "settings.json"
$logFilePath = "TAP_Generation_Log.txt"
$outpath = "Results.csv"  # Fichier de sortie pour les résultats

# Fonction pour écrire dans le fichier log
function Write-Log {
    param (
        [string]$message
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logMessage = "$timestamp - $message"
    Add-Content -Path $logFilePath -Value $logMessage
}

# Connexion à Microsoft Graph avec les scopes requis
if ((Get-MgContext) -eq $null) {
    try {
        Connect-MgGraph -Scopes "User.Read.All", "UserAuthenticationMethod.ReadWrite.All"
        Write-Log "Connexion réussie à Microsoft Graph avec les permissions requises."
    } catch {
        Write-Log "Erreur de connexion à Microsoft Graph : $_"
        [System.Windows.Forms.MessageBox]::Show("Erreur de connexion à Microsoft Graph. Vérifiez les permissions et la connexion réseau.", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
} else {
    Write-Log "Déjà connecté à Microsoft Graph."
}

# Fonction pour valider la colonne UserName dans le CSV
function Validate-CSV {
    param ($path)
    if (Test-Path -Path $path) {
        $csvContent = Import-Csv -Path $path
        if ($csvContent -and $csvContent[0].PSObject.Properties["UserName"]) {
            return $true
        } else {
            [System.Windows.Forms.MessageBox]::Show("Le fichier CSV doit contenir une colonne 'UserName'.", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return $false
        }
    } else {
        Write-Log "Chemin de fichier CSV invalide : $path"
        return $false
    }
}

# Charger les paramètres depuis le fichier JSON
$lastCSVPath = ""
if (Test-Path $settingsFilePath) {
    $settings = Get-Content -Path $settingsFilePath | ConvertFrom-Json
    if ($settings.LastCSVPath) {
        $lastCSVPath = $settings.LastCSVPath
    }
} else {
    $settings = @{ LastCSVPath = "" }
}

# Création de la fenêtre principale
$form = New-Object System.Windows.Forms.Form
$form.Text = "Génération de TAP"
$form.Size = New-Object System.Drawing.Size(500, 580)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Ajout d'un label de titre
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Outil de Génération de Temporary Access Pass (TAP)"
$titleLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
$titleLabel.AutoSize = $true
$form.Controls.Add($titleLabel)

# Ajout d'un PictureBox pour afficher le logo
if (Test-Path $logoPath) {
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Image = [System.Drawing.Image]::FromFile($logoPath)
    $pictureBox.SizeMode = 'Zoom'
    $pictureBox.Location = New-Object System.Drawing.Point(380, 10)
    $pictureBox.Size = New-Object System.Drawing.Size(100, 60)
    $form.Controls.Add($pictureBox)
}

# Label pour le chemin du fichier CSV
$labelPath = New-Object System.Windows.Forms.Label
$labelPath.Text = "Chemin du fichier CSV :"
$labelPath.Location = New-Object System.Drawing.Point(10, 100)
$labelPath.AutoSize = $true
$form.Controls.Add($labelPath)

# Champ pour afficher le chemin du fichier CSV sélectionné
$textBoxPath = New-Object System.Windows.Forms.TextBox
$textBoxPath.Location = New-Object System.Drawing.Point(10, 130)
$textBoxPath.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($textBoxPath)

# Bouton pour démarrer
$buttonStart = New-Object System.Windows.Forms.Button
$buttonStart.Text = "Démarrer"
$buttonStart.Location = New-Object System.Drawing.Point(10, 450)
$buttonStart.Size = New-Object System.Drawing.Size(100, 30)
$buttonStart.Enabled = $false  # Initialisé à false, sera activé si CSV valide
$form.Controls.Add($buttonStart)

# Si le dernier chemin CSV est valide, l'afficher et activer le bouton Démarrer
if (($lastCSVPath -ne "") -and (Validate-CSV -path $lastCSVPath)) {
    $textBoxPath.Text = $lastCSVPath
    $buttonStart.Enabled = $true
}

# Bouton pour parcourir et sélectionner le fichier CSV
$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = "Parcourir"
$buttonBrowse.Location = New-Object System.Drawing.Point(370, 130)
$buttonBrowse.Size = New-Object System.Drawing.Size(100, 25)
$form.Controls.Add($buttonBrowse)

# Barre de progression
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 170)
$progressBar.Size = New-Object System.Drawing.Size(460, 20)
$progressBar.Style = 'Continuous'
$form.Controls.Add($progressBar)

# Label de progression
$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Text = "Progression : 0%"
$progressLabel.Location = New-Object System.Drawing.Point(10, 195)
$progressLabel.AutoSize = $true
$form.Controls.Add($progressLabel)

# Zone de texte pour afficher les résultats en temps réel
$resultsBox = New-Object System.Windows.Forms.TextBox
$resultsBox.Location = New-Object System.Drawing.Point(10, 220)
$resultsBox.Size = New-Object System.Drawing.Size(460, 180)
$resultsBox.Multiline = $true
$resultsBox.ScrollBars = "Vertical"
$resultsBox.ReadOnly = $true
$form.Controls.Add($resultsBox)

# Label pour afficher le statut
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Statut : Prêt"
$statusLabel.Location = New-Object System.Drawing.Point(10, 410)
$statusLabel.AutoSize = $true
$form.Controls.Add($statusLabel)

$buttonCopyClipboard = New-Object System.Windows.Forms.Button
$buttonCopyClipboard.Text = "Copier dans le presse-papier"
$buttonCopyClipboard.Location = New-Object System.Drawing.Point(120, 450)
$buttonCopyClipboard.Size = New-Object System.Drawing.Size(160, 30)
$buttonCopyClipboard.Enabled = $false
$form.Controls.Add($buttonCopyClipboard)

$buttonOpenCSV = New-Object System.Windows.Forms.Button
$buttonOpenCSV.Text = "Ouvrir le fichier CSV"
$buttonOpenCSV.Location = New-Object System.Drawing.Point(290, 450)
$buttonOpenCSV.Size = New-Object System.Drawing.Size(100, 30)
$buttonOpenCSV.Enabled = $false
$form.Controls.Add($buttonOpenCSV)

$buttonNew = New-Object System.Windows.Forms.Button
$buttonNew.Text = "Nouveau"
$buttonNew.Location = New-Object System.Drawing.Point(400, 450)
$buttonNew.Size = New-Object System.Drawing.Size(70, 30)
$form.Controls.Add($buttonNew)

# Parcourir et sélectionner le fichier CSV
$buttonBrowse.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = "CSV Files|*.csv"
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxPath.Text = $fileDialog.FileName
        if (Validate-CSV -path $fileDialog.FileName) {
            $buttonStart.Enabled = $true
            $settings.LastCSVPath = $fileDialog.FileName
            $settings | ConvertTo-Json | Set-Content -Path $settingsFilePath
        }
    }
})

# Démarrer la génération de TAP
$buttonStart.Add_Click({
    $statusLabel.Text = "Statut : En cours..."
    Write-Log "Début de la génération de TAP"
    $buttonStart.Enabled = $false
    $buttonBrowse.Enabled = $false

    $properties = @{ isUsableOnce = $false } | ConvertTo-Json
    $hash = @{}
    $users = (Import-Csv -Path $textBoxPath.Text).UserName
    Write-Log "Source sélectionnée : Fichier CSV. Utilisateurs récupérés : $($users.Count)"

    # Initialisation de la barre de progression
    $totalUsers = $users.Count
    $progressBar.Maximum = $totalUsers
    $counter = 0

    # Boucle pour générer le TAP pour chaque utilisateur
    foreach ($user in $users) {
        try {
            $tapResponse = New-MgUserAuthenticationTemporaryAccessPassMethod -UserId $user -BodyParameter $properties
            $tap = $tapResponse.TemporaryAccessPass
            $hash.add($user, $tap)
            $resultsBox.AppendText("Succès : $user - TAP : $tap`n")
            Write-Log "Succès : $user - TAP : $tap"
        } catch {
            $hash.add($user, "Erreur")
            $resultsBox.AppendText("Erreur pour l'utilisateur $user : $_`n")
            Write-Log "Erreur pour l'utilisateur $user : $_"
        }
        $counter++
        $progressBar.Value = $counter
        $progressLabel.Text = "Progression : $([math]::Round(($counter / $totalUsers) * 100))%"
    }

    # Enregistrer les résultats dans un fichier CSV
    # Exporter le CSV avec guillemets

# Exporter le CSV avec les guillemets (comportement par défaut)
$hash.GetEnumerator() | Select-Object -Property @{N='UserPrincipalName';E={$_.Key}}, @{N='TemporaryAccessPass';E={$_.Value}} | Export-Csv -Path $outpath -NoTypeInformation -Encoding UTF8

# Supprimer uniquement les guillemets autour des champs, sans toucher aux guillemets internes
(Get-Content -Path $outpath) | ForEach-Object {
    $_ -replace '(^"|"$)', '' -replace '(","|",")', ',' 
} | Set-Content -Path $outpath -Encoding UTF8



Write-Log "Les résultats ont été enregistrés dans : $outpath"

    # Finaliser le statut et activer les boutons
    $statusLabel.Text = "Statut : Terminé"
    Write-Log "Génération de TAP terminée."
    $buttonCopyClipboard.Enabled = $true
    $buttonOpenCSV.Enabled = $true
    $buttonNew.Enabled = $true
})

# Copier dans le presse-papier
$buttonCopyClipboard.Add_Click({
    $clipboardText = "UserPrincipalName;TAP`n"
    foreach ($entry in $hash.GetEnumerator()) {
        $clipboardText += "$($entry.Key);$($entry.Value)`n"
    }
    [System.Windows.Forms.Clipboard]::SetText($clipboardText)
    [System.Windows.Forms.MessageBox]::Show("Les résultats ont été copiés dans le presse-papier.", "Copié", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Ouvrir le fichier CSV généré
$buttonOpenCSV.Add_Click({
    Invoke-Item (Resolve-Path -Path $outpath).Path
})

# Réinitialiser l'interface pour un nouveau traitement
$buttonNew.Add_Click({
    $textBoxPath.Clear()
    $resultsBox.Clear()
    $progressBar.Value = 0
    $progressLabel.Text = "Progression : 0%"
    $statusLabel.Text = "Statut : Prêt"
    $buttonStart.Enabled = $false
    $buttonCopyClipboard.Enabled = $false
    $buttonOpenCSV.Enabled = $false
    $buttonNew.Enabled = $false
    $hash.Clear()
    Write-Log "Interface réinitialisée pour un nouveau traitement."
})

# Afficher la fenêtre principale
$form.ShowDialog()
