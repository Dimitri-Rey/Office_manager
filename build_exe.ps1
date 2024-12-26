# Paramètres de compilation
$inputScript = "uninstall_office_gui.ps1"
$outputExe = "Office_Uninstaller.exe"
$iconPath = $null  # Vous pouvez ajouter une icône personnalisée si vous le souhaitez
$title = "Office Uninstaller"
$company = "Votre Entreprise"
$product = "Office Uninstaller"
$version = "1.0.0"
$copyright = "© $(Get-Date -Format yyyy)"

# Compilation en exe
Invoke-ps2exe -InputFile $inputScript `
              -OutputFile $outputExe `
              -IconFile $iconPath `
              -Title $title `
              -Company $company `
              -Product $product `
              -Version $version `
              -Copyright $copyright `
              -RequireAdmin `
              -NoConsole `
              -STA 