Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Gestionnaire Office 365" 
    Height="600" 
    Width="800" 
    WindowStartupLocation="CenterScreen"
    Background="#F0F0F0">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="15,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#106EBE"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#CCE4F4"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,8"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0">
            <TextBlock Text="Gestionnaire Office 365" 
                     FontSize="24" 
                     FontWeight="Bold" 
                     Margin="0,0,0,20"
                     Foreground="#0078D4"/>

            <GroupBox Header="Versions à désinstaller" Padding="10" Margin="0,0,0,10">
                <StackPanel>
                    <CheckBox x:Name="chkOffice365" Content="Microsoft 365 (Office 365)" IsChecked="True"/>
                    <CheckBox x:Name="chkOffice2021" Content="Office 2021" IsChecked="True"/>
                    <CheckBox x:Name="chkOffice2019" Content="Office 2019" IsChecked="True"/>
                    <CheckBox x:Name="chkOffice2016" Content="Office 2016" IsChecked="True"/>
                </StackPanel>
            </GroupBox>

            <GroupBox Header="Options" Padding="10" Margin="0,0,0,10">
                <StackPanel>
                    <CheckBox x:Name="chkCleanRegistry" Content="Nettoyer le registre Windows" IsChecked="True"/>
                    <CheckBox x:Name="chkRemoveFolders" Content="Supprimer les dossiers d'installation" IsChecked="True"/>
                </StackPanel>
            </GroupBox>

            <GroupBox Header="Installation" Padding="10">
                <StackPanel>
                    <CheckBox x:Name="chkReinstallO365" Content="Réinstaller Microsoft 365 (FR) après la désinstallation" IsChecked="False"/>
                </StackPanel>
            </GroupBox>
        </StackPanel>

        <Border Grid.Row="1" 
                Margin="0,10" 
                Background="White" 
                BorderBrush="#DDD" 
                BorderThickness="1" 
                CornerRadius="4">
            <ScrollViewer>
                <TextBox x:Name="txtLog" 
                         IsReadOnly="True" 
                         TextWrapping="Wrap"
                         Padding="10"
                         FontFamily="Consolas"
                         Background="Transparent"
                         BorderThickness="0"/>
            </ScrollViewer>
        </Border>

        <Grid Grid.Row="2" Margin="0,10,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <ProgressBar x:Name="progressBar" 
                        Height="4" 
                        Margin="0,0,10,0" 
                        Background="#F0F0F0"
                        Foreground="#0078D4"
                        BorderThickness="0"/>
            
            <Button x:Name="btnUninstall" 
                    Grid.Column="1" 
                    Content="EXÉCUTER" 
                    Margin="0,0,10,0"
                    Width="120"/>
            
            <Button x:Name="btnCancel" 
                    Grid.Column="2" 
                    Content="FERMER"
                    Width="120"
                    Background="#E81123"/>
        </Grid>
    </Grid>
</Window>
"@

# Création de la fenêtre
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Récupération des contrôles
$controls = @(
    'btnUninstall', 'btnCancel', 'txtLog', 'progressBar',
    'chkOffice365', 'chkOffice2021', 'chkOffice2019', 'chkOffice2016',
    'chkCleanRegistry', 'chkRemoveFolders', 'chkReinstallO365'
)

$controls | ForEach-Object {
    Set-Variable -Name $_ -Value $window.FindName($_)
}

# Fonction pour ajouter du texte au log avec timestamp
function Write-Log {
    param($Message, [switch]$Error)
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    $txtLog.Dispatcher.Invoke([Action]{
        if ($Error) {
            $txtLog.AppendText("❌ $logMessage`r`n")
        } else {
            $txtLog.AppendText("✓ $logMessage`r`n")
        }
        $txtLog.ScrollToEnd()
    })
}

# Fonction pour mettre à jour la barre de progression
function Update-Progress {
    param([int]$Value)
    $progressBar.Dispatcher.Invoke([Action]{
        $progressBar.Value = $Value
    })
}

# Fonction pour télécharger l'outil de déploiement Office
function Get-ODTTool {
    param($DestinationPath)
    
    try {
        Write-Log "Téléchargement de l'outil de déploiement Office..."
        
        # URL directe vers l'ODT (mise à jour régulièrement par Microsoft)
        $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16501-20196.exe"
        $odtPath = Join-Path $DestinationPath "ODT.exe"
        
        # Téléchargement avec .NET WebClient (plus fiable)
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($odtUrl, $odtPath)
        
        # Extraire setup.exe
        Write-Log "Extraction de setup.exe..."
        Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:$DestinationPath" -Wait
        
        # Vérifier que setup.exe existe
        $setupPath = Join-Path $DestinationPath "setup.exe"
        if (Test-Path $setupPath) {
            Write-Log "Outil de déploiement Office prêt"
            return $true
        } else {
            Write-Log "Échec de l'extraction de setup.exe" -Error
            return $false
        }
    }
    catch {
        Write-Log "Erreur lors du téléchargement : $($_.Exception.Message)" -Error
        return $false
    }
}

# Modifier la fonction Test-Prerequisites
function Test-Prerequisites {
    Write-Log "Vérification des prérequis..."
    
    try {
        # Vérifier l'espace disque
        $systemDrive = $env:SystemDrive
        $driveInfo = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='$systemDrive'"
        $freeSpaceGB = [math]::Round($driveInfo.FreeSpace / 1GB, 2)
        
        if ($freeSpaceGB -lt 5) {
            Write-Log "Espace disque insuffisant!" -Error
            [System.Windows.MessageBox]::Show(
                "L'espace disque disponible est insuffisant.`n`nVeuillez libérer au moins 5 Go d'espace disque.",
                "Espace insuffisant",
                "OK",
                "Warning"
            )
            return $false
        }
        
        # Vérifier la connexion Internet
        if (-not (Test-Connection -ComputerName "www.microsoft.com" -Count 1 -Quiet)) {
            Write-Log "Pas de connexion Internet!" -Error
            [System.Windows.MessageBox]::Show(
                "Une connexion Internet est requise pour télécharger les outils nécessaires.",
                "Erreur de connexion",
                "OK",
                "Error"
            )
            return $false
        }
        
        return $true
    }
    catch {
        Write-Log "Erreur lors de la vérification des prérequis : $($_.Exception.Message)" -Error
        return $false
    }
}

# Modifier la fonction Start-OfficeUninstall
function Start-OfficeUninstall {
    try {
        if (-not (Test-Prerequisites)) { 
            $btnUninstall.IsEnabled = $true
            return 
        }
        
        Update-Progress 10
        
        # Créer un dossier temporaire unique
        $tempFolder = Join-Path $env:TEMP "OfficeUninstall_$(Get-Random)"
        New-Item -ItemType Directory -Force -Path $tempFolder | Out-Null
        Write-Log "Dossier temporaire créé : $tempFolder"
        
        Update-Progress 20
        
        # Télécharger et extraire l'outil
        if (-not (Get-ODTTool -DestinationPath $tempFolder)) {
            throw "Impossible de préparer l'outil de déploiement Office"
        }
        
        Update-Progress 30
        
        # Configuration de désinstallation
        $configXml = @"
<Configuration>
    <Display Level="None" />
    <Remove All="True" />
</Configuration>
"@
        $configXml | Out-File "$tempFolder\config.xml" -Encoding UTF8
        Write-Log "Configuration de désinstallation préparée"
        
        Update-Progress 40
        
        # Désinstallation
        Write-Log "Désinstallation d'Office en cours... (cette étape peut prendre plusieurs minutes)"
        Start-Process -FilePath "$tempFolder\setup.exe" -ArgumentList "/configure $tempFolder\config.xml" -Wait
        Write-Log "Désinstallation terminée"
        
        Update-Progress 60
        
        # Nettoyage du registre
        if ($chkCleanRegistry.IsChecked) {
            Write-Log "Nettoyage du registre..."
            $regPaths = @(
                "HKLM:\SOFTWARE\Microsoft\Office",
                "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office"
            )
            foreach ($path in $regPaths) {
                if (Test-Path $path) {
                    Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Log "Registre nettoyé : $path"
                }
            }
        }
        
        Update-Progress 70
        
        # Nettoyage des dossiers
        if ($chkRemoveFolders.IsChecked) {
            Write-Log "Suppression des dossiers..."
            $officePaths = @(
                "C:\Program Files\Microsoft Office",
                "C:\Program Files (x86)\Microsoft Office"
            )
            foreach ($path in $officePaths) {
                if (Test-Path $path) {
                    Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Log "Dossier supprimé : $path"
                }
            }
        }
        
        Update-Progress 80
        
        # Réinstallation si demandée
        if ($chkReinstallO365.IsChecked) {
            Write-Log "Préparation de la réinstallation..."
            Install-Office365FR
        }
        
        # Nettoyage final
        Write-Log "Nettoyage des fichiers temporaires..."
        Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        
        Update-Progress 100
        Write-Log "Opération terminée avec succès!"
        
        [System.Windows.MessageBox]::Show(
            "Toutes les opérations ont été effectuées avec succès.",
            "Terminé",
            "OK",
            "Information"
        )
    }
    catch {
        Write-Log $_.Exception.Message -Error
        [System.Windows.MessageBox]::Show(
            "Une erreur est survenue pendant l'opération.`n`nDétails: $($_.Exception.Message)",
            "Erreur",
            "OK",
            "Error"
        )
    }
    finally {
        # S'assurer que le dossier temporaire est supprimé
        if (Test-Path $tempFolder) {
            Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
        Update-Progress 0
        $btnUninstall.IsEnabled = $true
    }
}

# Fonction d'installation Office 365 FR
function Install-Office365FR {
    try {
        Write-Log "Configuration de l'installation Office 365 FR"
        
        $configXml = @"
<Configuration>
    <Add OfficeClientEdition="64" Channel="Current">
        <Product ID="O365ProPlusRetail">
            <Language ID="fr-fr" />
            <ExcludeApp ID="Groove" />
            <ExcludeApp ID="Lync" />
        </Product>
    </Add>
    <Display Level="None" AcceptEULA="TRUE" />
    <Property Name="AUTOACTIVATE" Value="1" />
</Configuration>
"@
        $configXml | Out-File "$tempFolder\configInstall.xml" -Encoding UTF8
        
        Write-Log "Installation d'Office 365 FR en cours... (cette étape peut prendre plusieurs minutes)"
        Start-Process -FilePath "$tempFolder\setup.exe" -ArgumentList "/configure $tempFolder\configInstall.xml" -Wait
        Write-Log "Installation d'Office 365 FR terminée"
    }
    catch {
        Write-Log "Erreur lors de l'installation : $($_.Exception.Message)" -Error
        throw
    }
}

# Événements des boutons
$btnUninstall.Add_Click({
    $btnUninstall.IsEnabled = $false
    $txtLog.Clear()
    Start-OfficeUninstall
})

$btnCancel.Add_Click({
    $window.Close()
})

# Vérification des droits administrateur
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    [System.Windows.MessageBox]::Show(
        "Ce programme doit être exécuté en tant qu'administrateur.",
        "Droits insuffisants",
        "OK",
        "Error"
    )
    exit
}

# Affichage de la fenêtre
$window.ShowDialog() | Out-Null 