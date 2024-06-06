# Définition de la fonction pour trouver l'UO enfant en fonction du type d'utilisateur
function Get-ChildOU {
    param([string]$childOUName)
    
    $rootOU = "OU=ActiveUsers,OU=Domain Controllers,DC=redlab,DC=local"  # Remplacer par le DN de votre OU racine
    $ou = Get-ADOrganizationalUnit -Filter "Name -eq '$childOUName'" -SearchBase $rootOU
    
    if ($ou) {
        return $ou
    } else {
        Write-Error "L'UO enfant '$childOUName' n'a pas été trouvée sous '$rootOU'."
        return $null
    }
}

# Sélection du fichier .txt à traiter
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.InitialDirectory = [System.Environment]::GetFolderPath('Desktop')
$openFileDialog.Filter = "Fichiers texte (*.txt)|*.txt"
$openFileDialog.Title = "Sélectionner le fichier texte des utilisateurs"

if ($openFileDialog.ShowDialog() -eq 'OK') {
    $file = $openFileDialog.FileName
} else {
    Write-Host "Aucun fichier sélectionné. Le script va s'arrêter."
    exit
}

# Définir le chemin du fichier de sortie sur le bureau
$outFile = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'login.txt')

# Création d'un tableau pour stocker les informations des utilisateurs créés
$createdUsers = @()

# Lecture du fichier .txt et création des utilisateurs
$users = Get-Content $file

foreach ($userLine in $users) {
    # Ignorer les lignes commentées
    if ($userLine -match '^\s*#') {
        continue
    }
    
    # Ignorer les lignes vides
    if ([string]::IsNullOrWhiteSpace($userLine)) {
        continue
    }

    # Supprimer le ';' à la fin de la ligne
    $userLine = $userLine.TrimEnd(';')

    $userInfo = $userLine -split ","
    
    if ($userInfo.Count -ne 4) {
        Write-Host "Le format de ligne est incorrect pour : $userLine"
        continue
    }
    
    $nom = $userInfo[0].Trim()
    $prenom = $userInfo[1].Trim()
    $dob = $userInfo[2].Trim()
    $childOUName = $userInfo[3].Trim()
    
    # Générer le nom d'utilisateur
    $username = ($prenom.Substring(0,1) + "." + $nom).ToLower()
    
    # Générer le mot de passe
    $initialPassword = ($prenom.Substring(0,1) + $nom.Substring(0,1) + $dob + "*").ToLower()
    $securePassword = ConvertTo-SecureString $initialPassword -AsPlainText -Force
    
    # Trouver l'UO enfant appropriée
    $childOU = Get-ChildOU -childOUName $childOUName
    if (-not $childOU) {
        Write-Host "Impossible de créer l'utilisateur $prenom $nom. UO enfant '$childOUName' introuvable."
        continue
    }
    
    # Créer l'utilisateur
    $params = @{
        'Name'              = "$prenom $nom"
        'SamAccountName'    = $username
        'UserPrincipalName' = "$username@redlab.local"
        'GivenName'         = $prenom
        'Surname'           = $nom
        'Enabled'           = $true
        'Path'              = $childOU.DistinguishedName
        'AccountPassword'   = $securePassword
        'ChangePasswordAtLogon' = $true
        'CannotChangePassword' = $false
        'PasswordNeverExpires' = $false
    }
    
    try {
        $userObj = New-ADUser @params -ErrorAction Stop
        
        # Ajouter les informations à la liste des utilisateurs créés
        $userInfo = "$prenom $nom, $initialPassword"
        $createdUsers += $userInfo
        
        Write-Host "L'utilisateur $prenom $nom a été créé avec succès dans $($childOU.Name). Mot de passe initial : $initialPassword"
    } catch {
        Write-Host "Erreur lors de la création de l'utilisateur $prenom $nom : $_"
    }
}

# Écrire les informations des utilisateurs dans le fichier de sortie
try {
    $createdUsers | Out-File -FilePath $outFile -Encoding UTF8
    Write-Host "Les informations des utilisateurs ont été écrites dans $outFile."
} catch {
    Write-Host "Erreur lors de l'écriture dans le fichier $outFile : $_"
}

Write-Host "Le script a terminé l'importation des utilisateurs depuis $file."
