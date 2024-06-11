# Function to get or create a child OU
function Get-ChildOU {
    param([string]$childOUName, [string]$rootOU)
    
    $ou = Get-ADOrganizationalUnit -Filter "Name -eq '$childOUName'" -SearchBase $rootOU
    
    if ($ou) {
        return $ou
    } else {
        Write-Host "Child OU '$childOUName' not found under '$rootOU'. Creating it..."
        try {
            $newOU = New-ADOrganizationalUnit -Name $childOUName -Path $rootOU -ErrorAction Stop
            Write-Host "Child OU '$childOUName' created successfully under '$rootOU'."
            return $newOU
        } catch {
            Write-Error "Failed to create child OU '$childOUName': $($_)"
            return $null
        }
    }
}

# Function to select a text file
function Select-TextFile {
    Add-Type -AssemblyName System.Windows.Forms

    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = [System.Environment]::GetFolderPath('Desktop')
    $openFileDialog.Filter = "Text Files (*.txt)|*.txt"
    $openFileDialog.Title = "Select the user text file"

    if ($openFileDialog.ShowDialog() -eq 'OK') {
        return $openFileDialog.FileName
    } else {
        Write-Host "No file selected. The script will stop."
        exit
    }
}

# Function to create users
function Create-Users {
    param(
        [string]$file,
        [string]$rootOU
    )

    $users = Get-Content $file
    $createdUsers = @()

    $childOUs = @()

    # First pass to check and create OUs if necessary
    foreach ($userLine in $users) {
        # Ignore lines starting with #
        if ($userLine -match '^\s*#') {
            continue
        }
        
        if ([string]::IsNullOrWhiteSpace($userLine)) {
            continue
        }
        
        $userLine = $userLine.TrimEnd(';')
        $userInfo = $userLine -split ","
        
        if ($userInfo.Count -ne 4) {
            Write-Host "Incorrect line format for: $userLine"
            continue
        }
        
        $childOUName = $userInfo[3].Trim()
        
        if (-not $childOUs.Contains($childOUName)) {
            $childOU = Get-ChildOU -childOUName $childOUName -rootOU $rootOU
            if ($childOU) {
                $childOUs += $childOUName
            }
        }
    }

    # Wait for 5 seconds before creating users
    Write-Host "Waiting 5 seconds for AD synchronization..."
    for ($i = 5; $i -gt 0; $i--) {
        Write-Host "$i..."
        Start-Sleep -Seconds 1
    }

    # Second pass to create users
    foreach ($userLine in $users) {
        # Ignore lines starting with #
        if ($userLine -match '^\s*#') {
            continue
        }
        
        if ([string]::IsNullOrWhiteSpace($userLine)) {
            continue
        }
        
        $userLine = $userLine.TrimEnd(';')
        $userInfo = $userLine -split ","
        
        if ($userInfo.Count -ne 4) {
            Write-Host "Incorrect line format for: $userLine"
            continue
        }
        
        $nom = $userInfo[0].Trim()
        $prenom = $userInfo[1].Trim()
        $dob = $userInfo[2].Trim()
        $childOUName = $userInfo[3].Trim()
        
        # Generate username
        $username = ($prenom.Substring(0,1) + "." + $nom).ToLower()
        
        # Generate temporary password
        $initialPassword = ($prenom.Substring(0,1) + $nom.Substring(0,1) + $dob + "*").ToLower()
        $securePassword = ConvertTo-SecureString $initialPassword -AsPlainText -Force
        
        # Find Organizational Unit
        $childOU = Get-ADOrganizationalUnit -Filter "Name -eq '$childOUName'" -SearchBase $rootOU
        if (-not $childOU) {
            Write-Host "Cannot create user $prenom $nom. Child OU '$childOUName' not found."
            continue
        }
        
        # Create user
        $params = @{
            'Name'              = "$prenom $nom"
            'SamAccountName'    = $username
            'UserPrincipalName' = "$username@domain.local" # Change the domain.local to your DN
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
            
            Write-Host "User $prenom $nom created successfully in $($childOU.Name). Initial password: $initialPassword"
        } catch {
            Write-Host "Error creating user $prenom $nom : $($_). Retrying after AD sync..."
            Start-Sleep -Seconds 5
            try {
                $userObj = New-ADUser @params -ErrorAction Stop
                
                $userInfo = "$prenom $nom, $username, $initialPassword"
                $createdUsers += $userInfo
                
                Write-Host "User $prenom $nom created successfully in $($childOU.Name) after AD sync. Initial password: $initialPassword"
            } catch {
                Write-Host "Failed again to create user $prenom $nom : $($_)"
            }
        }
    }

    return $createdUsers
}

# Function to write user information to a file
function Write-UserInformationToFile {
    param(
        [array]$createdUsers,
        [string]$outFile
    )

    try {
        $createdUsers | Out-File -FilePath $outFile -Encoding UTF8
        Write-Host "User information written to $outFile."
    } catch {
        Write-Host "Error writing to file $outFile : $($_)"
    }
}

# Main script execution
try {
    $rootOU = "OU=Domain Controllers,DC=redlab,DC=local"  # Replace with your root OU's DN
    
    # Select .txt file
    $file = Select-TextFile
    
    if (-not $file) {
        exit
    }

    # Create users
    $createdUsers = Create-Users -file $file -rootOU $rootOU

    # Output file with username and temporary password
    $outFile = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'login.txt')

    # Write user information to file
    Write-UserInformationToFile -createdUsers $createdUsers -outFile $outFile

    Write-Host "User import from $file completed."
} catch {
    Write-Host "Error: $($_)"
}
