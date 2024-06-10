function Get-ChildOU {
    param([string]$childOUName)
    
    $rootOU = "OU=Domain Controllers,DC=domain,DC=local"  # Remplace by the DN of you're root OU
    $ou = Get-ADOrganizationalUnit -Filter "Name -eq '$childOUName'" -SearchBase $rootOU
    
    if ($ou) {
        return $ou
    } else {
        Write-Error "L'UO enfant '$childOUName' n'a pas été trouvée sous '$rootOU'."
        return $null
    }
}

# Select .txt
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

# output file with username and temporary password
$outFile = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'login.txt')

$createdUsers = @()

$users = Get-Content $file

foreach ($userLine in $users) {

    # line who start with # is ignored
    if ($userLine -match '^\s*#') {
        continue
    }
    
    if ([string]::IsNullOrWhiteSpace($userLine)) {
        continue
    }
    
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
    
    # Genrate username
    $username = ($prenom.Substring(0,1) + "." + $nom).ToLower()
    
    # Generate temporary password
    $initialPassword = ($prenom.Substring(0,1) + $nom.Substring(0,1) + $dob + "*").ToLower()
    $securePassword = ConvertTo-SecureString $initialPassword -AsPlainText -Force
    
    # Find Orginastion Unite
    $childOU = Get-ChildOU -childOUName $childOUName
    if (-not $childOU) {
        Write-Host "Impossible de créer l'utilisateur $prenom $nom. UO enfant '$childOUName' introuvable."
        continue
    }
    
    # Create user
    $params = @{
        'Name'              = "$prenom $nom"
        'SamAccountName'    = $username
        'UserPrincipalName' = "$username@domain.local" #change the domain.local by you're DN
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
        
        $userInfo = "$prenom $nom, $username, $initialPassword"
        $createdUsers += $userInfo
        
        Write-Host "L'utilisateur $prenom $nom a été créé avec succès dans $($childOU.Name). Mot de passe initial : $initialPassword"
    } catch {
        Write-Host "Erreur lors de la création de l'utilisateur $prenom $nom : $_"
    }
}

# Write output file
try {
    $createdUsers | Out-File -FilePath $outFile -Encoding UTF8
    Write-Host "Les informations des utilisateurs ont été écrites dans $outFile."
} catch {
    Write-Host "Erreur lors de l'écriture dans le fichier $outFile : $_"
}

Write-Host "Le script a terminé l'importation des utilisateurs depuis $file."
