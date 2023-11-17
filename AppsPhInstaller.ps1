Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to create or update registry entry
function Set-RegistryEntry ($keyPath, $valueName, $value) {
    $registryKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($keyPath, $true)

    if (!$registryKey) {
        $null = [Microsoft.Win32.Registry]::LocalMachine.CreateSubKey($keyPath)
        $registryKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($keyPath, $true)
    }

    $registryKey.SetValue($valueName, $value, [Microsoft.Win32.RegistryValueKind]::String)
    $registryKey.Close()
}

# Function to copy files from source to destination
function Copy-Files ($sourcePath, $destinationPath) {
    $sourceFiles = Get-ChildItem -Path $sourcePath -File
    foreach ($file in $sourceFiles) {
        $destinationFile = Join-Path $destinationPath $file.Name
        Copy-Item -Path $file.FullName -Destination $destinationFile -Force
    }
}

# Function to add exclusion path to Windows Defender
function Add-WindowsDefenderExclusion ($path) {
    Add-MpPreference -ExclusionPath $path
}

# Function to set full control permissions on a folder
function Set-FullControlPermissions ($folderPath) {
    $acl = Get-Acl -Path $folderPath
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule("Users", "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
    $acl.SetAccessRule($rule)
    Set-Acl -Path $folderPath -AclObject $acl
}

# Function to create a folder in C:\ and set permissions
function Create-Folder ($folderName) {
    $folderPath = "C:\$folderName"
    
    if (-not (Test-Path -Path $folderPath)) {
        New-Item -Path $folderPath -ItemType Directory -Force
        Set-FullControlPermissions -folderPath $folderPath
    }

    return $folderPath
}

# Function to get the directory path for "Obtener Directorio APPS"
function Get-AppsDirectoryPath {
    return Join-Path -Path $env:USERPROFILE -ChildPath "AppData\Local\Apps\2.0"
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Instalador Apps PH"
$form.Size = New-Object System.Drawing.Size(950, 400)

# Create buttons
$buttonCRMPH = New-Object System.Windows.Forms.Button
$buttonCRMPH.Text = "Instalar CRMPH Portable"
$buttonCRMPH.Size = New-Object System.Drawing.Size(150, 65)
$buttonCRMPH.Location = New-Object System.Drawing.Point(70, 65)

# Create button to get "Obtener Directorio APPS"
$buttonGetAppsDirectory = New-Object System.Windows.Forms.Button
$buttonGetAppsDirectory.Text = "Obtener Directorio APPS"
$buttonGetAppsDirectory.Size = New-Object System.Drawing.Size(180, 65)
$buttonGetAppsDirectory.Location = New-Object System.Drawing.Point(230, 65)

# Create label for directory path
$labelAppsDirectory = New-Object System.Windows.Forms.Label
$labelAppsDirectory.Size = New-Object System.Drawing.Size(600, 20)
$labelAppsDirectory.Location = New-Object System.Drawing.Point(230, 135)
$labelAppsDirectory.Text = ""

# Create button to remove "2.0"
$buttonRemoveTwoDotZero = New-Object System.Windows.Forms.Button
$buttonRemoveTwoDotZero.Text = "Eliminar 2.0"
$buttonRemoveTwoDotZero.Size = New-Object System.Drawing.Size(115, 65)
$buttonRemoveTwoDotZero.Location = New-Object System.Drawing.Point(420, 65)
$buttonRemoveTwoDotZero.Enabled = $false  # Initially disabled

# Create button to kill "ClickOnce" process
$buttonKillClickOnce = New-Object System.Windows.Forms.Button
$buttonKillClickOnce.Text = "Kill ClickOnce"
$buttonKillClickOnce.Size = New-Object System.Drawing.Size(115, 65)
$buttonKillClickOnce.Location = New-Object System.Drawing.Point(560, 65)

# Create button to open Uninstall Programs window
$buttonUninstallApps = New-Object System.Windows.Forms.Button
$buttonUninstallApps.Text = "Desinstalar Apps"
$buttonUninstallApps.Size = New-Object System.Drawing.Size(180, 65)
$buttonUninstallApps.Location = New-Object System.Drawing.Point(690, 65)

# Create button to install CRMPH Azure
$buttonInstallCRMPHAzure = New-Object System.Windows.Forms.Button
$buttonInstallCRMPHAzure.Text = "Instalar CRMPH Azure"
$buttonInstallCRMPHAzure.Size = New-Object System.Drawing.Size(180, 65)
$buttonInstallCRMPHAzure.Location = New-Object System.Drawing.Point(70, 150)

# Event handler for CRMPH Azure button click
$buttonInstallCRMPHAzure.Add_Click({
    $CRMUrl = "https://uscldhphwap01.azurewebsites.net/Install/CRM/CRMPH.application"
    Start-Process -FilePath "rundll32.exe" -ArgumentList "dfshim.dll,ShOpenVerbApplication $CRMUrl"
})

# Create button to install CRMPH 02
$buttonInstallCRMPH02 = New-Object System.Windows.Forms.Button
$buttonInstallCRMPH02.Text = "Instalar CRMPH 02"
$buttonInstallCRMPH02.Size = New-Object System.Drawing.Size(180, 65)
$buttonInstallCRMPH02.Location = New-Object System.Drawing.Point(260, 150)

# Event handler for CRMPH 02 button click
$buttonInstallCRMPH02.Add_Click({
    $CRMUrl = "\\Mxoccaph02\actual\CRMPH.application"
    Start-Process -FilePath "rundll32.exe" -ArgumentList "dfshim.dll,ShOpenVerbApplication $CRMUrl"
})

# Event handler for CRMPH button click
$buttonCRMPH.Add_Click({
    # Create "CRMPH" folder in C:\
    $path = Create-Folder -folderName "CRMPH"

    # Set registry entry
    Set-RegistryEntry "SOFTWARE\Patrimonio hoy\CRMPH" "Directorio" $path

    # Copy files from "\\M02\crmph_portable" to "CRMPH" folder
    $sourcePath = "\\M02\crmph_portable"
    Copy-Files -sourcePath $sourcePath -destinationPath $path

    # Add exclusion path to Windows Defender
    Add-WindowsDefenderExclusion -path $path
})

# Event handler for "Obtener Directorio APPS" button click
$buttonGetAppsDirectory.Add_Click({
    # Get directory path for "Obtener Directorio APPS"
    $directoryPath = Get-AppsDirectoryPath

    # Set the label text to the obtained directory path
    $labelAppsDirectory.Text = $directoryPath

    # Enable the "Eliminar 2.0" button
    $buttonRemoveTwoDotZero.Enabled = $true
})

# Event handler for "Eliminar 2.0" button click
$buttonRemoveTwoDotZero.Add_Click({
    # Remove all files from the obtained directory path
    $directoryPath = Get-AppsDirectoryPath
    Remove-Item -Path "$directoryPath\*" -Recurse -Force
})

# Event handler for "Kill ClickOnce" button click
$buttonKillClickOnce.Add_Click({
    # Kill all processes with name "dfsvc.exe" (ClickOnce)
    Get-Process -Name "dfsvc" | ForEach-Object { Stop-Process -Id $_.Id -Force }
})

# Event handler for "Desinstalar Apps" button click
$buttonUninstallApps.Add_Click({
    # Open Uninstall Programs window in Control Panel
    Start-Process "control.exe" -ArgumentList "appwiz.cpl"
})

# Add controls to the form
$form.Controls.Add($buttonCRMPH)
$form.Controls.Add($buttonGetAppsDirectory)
$form.Controls.Add($labelAppsDirectory)
$form.Controls.Add($buttonRemoveTwoDotZero)
$form.Controls.Add($buttonKillClickOnce)
$form.Controls.Add($buttonUninstallApps)
$form.Controls.Add($buttonInstallCRMPHAzure)
$form.Controls.Add($buttonInstallCRMPH02)

# Show the form
[Windows.Forms.Application]::Run($form)

# Execute CRMPH.exe from "CRMPH" Path
$directoryPath = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Patrimonio hoy\CRMPH" -Name "Directorio" | Select-Object -ExpandProperty Directorio
$crmphExePath = Join-Path $directoryPath "CRMPH.exe"

if (Test-Path -Path $crmphExePath) {
    Start-Process -FilePath $crmphExePath
} else {
    Write-Host "CRMPH.exe not found in the specified directory."
}
