##################################################
##                                              ##
##    Script pour désinstaller ou installer     ##
##        OFFICE et installer des applis        ##
##                                              ##
##################################################
# OELIS RELAVE Thomas
# 10/12/2024
# V2.1

# Applis installables --> Adobe, Chrome, Firefox, 7-zip, Java, PDFcreator, Zoom

<# Comment cela fonctionne
    Il faut avant tout placer le fichier "Appli_script" 
    dans le C: pour pouvoir installer les logiciels.
    Ensuite, il suffira de lancer le script en tant qu'administrateur.
    Taper la commande --> Set-ExecutionPolicy RemoteSigned # Cliquez sur "oui pour tous"
#>

#Requires -Version 5.1

Add-Type -AssemblyName PresentationFramework

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ConfigurationXMLFile,
    [string]$NbOffice,
    [string]$Appli,
    [string]$OfficeInstallDownloadPath = "$env:TEMP\office365Install",
    [Switch]$Restart = $false
)

function Show-Message {
    param (
        [string]$message
    )
    [System.Windows.MessageBox]::Show($message, "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
}

function Create-Window {
    $window = New-Object System.Windows.Window
    $window.Title = "Office Installer"
    $window.SizeToContent = "WidthAndHeight"
    $window.WindowStartupLocation = "CenterScreen"

    $grid = New-Object System.Windows.Controls.Grid
    $window.Content = $grid

    $label = New-Object System.Windows.Controls.Label
    $label.Content = "Voulez-vous installer ou désinstaller Office ?"
    $label.HorizontalAlignment = "Center"
    $label.Margin = "10"
    $grid.Children.Add($label)

    $buttonInstall = New-Object System.Windows.Controls.Button
    $buttonInstall.Content = "Installer"
    $buttonInstall.Margin = "10"
    $buttonInstall.HorizontalAlignment = "Center"
    $buttonInstall.VerticalAlignment = "Center"
    $buttonInstall.Add_Click({ Execute-Action "Install" })
    $grid.Children.Add($buttonInstall)

    $buttonUninstall = New-Object System.Windows.Controls.Button
    $buttonUninstall.Content = "Désinstaller"
    $buttonUninstall.Margin = "10"
    $buttonUninstall.HorizontalAlignment = "Center"
    $buttonUninstall.VerticalAlignment = "Center"
    $buttonUninstall.Add_Click({ Execute-Action "Uninstall" })
    $grid.Children.Add($buttonUninstall)

    return $window
}

function Execute-Action {
    param (
        [string]$action
    )

    if ($action -eq "Install") {
        $NbOffice = "Install"
    } elseif ($action -eq "Uninstall") {
        $NbOffice = "Uninstall"
    }

    # Créer le dossier d'installation si nécessaire
    if (-not (Test-Path $OfficeInstallDownloadPath)) {
        New-Item -Path $OfficeInstallDownloadPath -ItemType Directory | Out-Null
    }

    # Télécharger l'ODT
    $ODTInstallLink = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe"
    Invoke-WebRequest -Uri $ODTInstallLink -OutFile (Join-Path $OfficeInstallDownloadPath 'ODTSetup.exe')

    # Exécuter l'ODT
    Start-Process (Join-Path $OfficeInstallDownloadPath 'ODTSetup.exe') -ArgumentList '/quiet /extract:$OfficeInstallDownloadPath' -Wait

    if ($NbOffice -eq "Install") {
        # Créer le fichier XML d'installation
        $OfficeXMLInstall = [XML]@"
        <Configuration>
            <Add OfficeClientEdition="64" Channel="Current">
                <Product ID="O365BusinessRetail">
                    <Language ID="fr-fr" />
                    <ExcludeApp ID="Groove" />
                    <ExcludeApp ID="Lync" />
                </Product>
            </Add>
            <Updates Enabled="TRUE" />
            <RemoveMSI />
        </Configuration>
        "@
        $OfficeXMLInstall.Save((Join-Path $OfficeInstallDownloadPath 'OfficeInstall.xml'))

        # Installer Office
        Start-Process (Join-Path $OfficeInstallDownloadPath 'Setup.exe') -ArgumentList "/configure (Join-Path $OfficeInstallDownloadPath 'OfficeInstall.xml')" -Wait -Verb RunAs
    } elseif ($NbOffice -eq "Uninstall") {
        # Créer le fichier XML de désinstallation
        $OfficeXMLUninstall = [XML]@"
        <Configuration>
            <Remove All="TRUE"/>
            <Display Level="None" AcceptEULA="TRUE"/>
        </Configuration>
        "@
        $OfficeXMLUninstall.Save((Join-Path $OfficeInstallDownloadPath 'OfficeUninstall.xml'))

        # Désinstaller Office
        Start-Process (Join-Path $OfficeInstallDownloadPath 'Setup.exe') -ArgumentList "/configure (Join-Path $OfficeInstallDownloadPath 'OfficeUninstall.xml')" -Wait -Verb RunAs
    }

    # Installer d'autres applications
    Install-Applications

    # Nettoyer les fichiers temporaires
    Remove-Item -Path $OfficeInstallDownloadPath -Recurse -Force
}

function Install-Applications {
    $applications = @('Google.Chrome', 'Mozilla.Firefox', 'Adobe.Acrobat.Reader.64-bit')
    foreach ($app in $applications) {
        $response = Read-Host "Voulez-vous installer $app ? (y/n)"
        if ($response -eq 'y' -or $response -eq 'Y') {
            winget install "$app" --accept-package-agreements
            Show-Message "$app a été installé."
        }
    }
}

# Vérification des droits administrateur
if (-not [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent().IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Show-Message "Access Denied. Please run with Administrator privileges."
    exit 1
}

# Créer et afficher la fenêtre
$window = Create-Window
$window.ShowDialog() | Out-Null 