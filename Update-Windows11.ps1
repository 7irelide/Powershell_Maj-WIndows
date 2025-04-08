#requires -Version 5.1
#requires -RunAsAdministrator

<#
.SYNOPSIS
    Script de mise à jour complète pour Windows 11.
.DESCRIPTION
    Ce script recherche et installe toutes les mises à jour Windows 11 disponibles,
    y compris les mises à jour de sécurité, cumulatives, facultatives, et les pilotes.
    Il génère un rapport détaillé des mises à jour installées et de leur statut.
.NOTES
    Auteur: AI Assistant
    Date: $(Get-Date -Format "dd/MM/yyyy")
    Prérequis: 
        - Windows 11
        - Droits administrateur
        - Connexion Internet active
#>

# Variables pour le rapport
$updateReport = @{
    "WindowsUpdates" = @{
        "Installed" = @()
        "Failed" = @()
    }
    "Drivers" = @{
        "Installed" = @()
        "Failed" = @()
    }
    "RestartRequired" = $false
}

# Fonction pour afficher les messages avec des couleurs
function Write-ColorMessage {
    param (
        [string]$Message,
        [string]$ForegroundColor = "White"
    )
    Write-Host $Message -ForegroundColor $ForegroundColor
}

# Fonction pour vérifier si un module est installé
function Test-ModuleInstalled {
    param ([string]$ModuleName)
    
    return (Get-Module -ListAvailable -Name $ModuleName) -ne $null
}

# Fonction pour installer un module si nécessaire
function Install-ModuleIfNotExists {
    param ([string]$ModuleName)
    
    if (-not (Test-ModuleInstalled -ModuleName $ModuleName)) {
        Write-ColorMessage "Le module $ModuleName n'est pas installé. Installation en cours..." "Yellow"
        try {
            Install-Module -Name $ModuleName -Force -Scope CurrentUser -ErrorAction Stop
            Write-ColorMessage "Module $ModuleName installé avec succès." "Green"
            Import-Module -Name $ModuleName -Force
            return $true
        }
        catch {
            Write-ColorMessage "Erreur lors de l'installation du module $ModuleName : $_" "Red"
            return $false
        }
    }
    else {
        Write-ColorMessage "Le module $ModuleName est déjà installé." "Green"
        Import-Module -Name $ModuleName -Force
        return $true
    }
}

# Fonction pour générer un rapport formaté
function Show-UpdateReport {
    Write-ColorMessage "`n===== RAPPORT DE FIN D'EXÉCUTION =====" "Cyan"
    
    Write-ColorMessage "`n[MISES À JOUR WINDOWS]" "Yellow"
    Write-ColorMessage "Mises à jour installées ($($updateReport.WindowsUpdates.Installed.Count)) :" "Green"
    if ($updateReport.WindowsUpdates.Installed.Count -eq 0) {
        Write-ColorMessage "  - Aucune mise à jour installée" "White"
    }
    else {
        foreach ($update in $updateReport.WindowsUpdates.Installed) {
            Write-ColorMessage "  - $update" "White"
        }
    }
    
    Write-ColorMessage "`nMises à jour ayant échoué ($($updateReport.WindowsUpdates.Failed.Count)) :" "Red"
    if ($updateReport.WindowsUpdates.Failed.Count -eq 0) {
        Write-ColorMessage "  - Aucun échec" "White"
    }
    else {
        foreach ($update in $updateReport.WindowsUpdates.Failed) {
            Write-ColorMessage "  - $update" "White"
        }
    }
    
    Write-ColorMessage "`n[PILOTES]" "Yellow"
    Write-ColorMessage "Pilotes installés ($($updateReport.Drivers.Installed.Count)) :" "Green"
    if ($updateReport.Drivers.Installed.Count -eq 0) {
        Write-ColorMessage "  - Aucun pilote installé" "White"
    }
    else {
        foreach ($driver in $updateReport.Drivers.Installed) {
            Write-ColorMessage "  - $driver" "White"
        }
    }
    
    Write-ColorMessage "`nPilotes ayant échoué ($($updateReport.Drivers.Failed.Count)) :" "Red"
    if ($updateReport.Drivers.Failed.Count -eq 0) {
        Write-ColorMessage "  - Aucun échec" "White"
    }
    else {
        foreach ($driver in $updateReport.Drivers.Failed) {
            Write-ColorMessage "  - $driver" "White"
        }
    }
    
    Write-ColorMessage "`n[ÉTAT DU SYSTÈME]" "Yellow"
    if ($updateReport.RestartRequired) {
        Write-ColorMessage "Un redémarrage est requis pour finaliser les mises à jour." "Red"
    }
    else {
        Write-ColorMessage "Aucun redémarrage n'est nécessaire." "Green"
    }
    
    Write-ColorMessage "`n=====================================" "Cyan"
}

# Vérification du système d'exploitation
Write-ColorMessage "Vérification du système d'exploitation..." "Cyan"
$osInfo = Get-CimInstance Win32_OperatingSystem
$osName = $osInfo.Caption
$osVersion = $osInfo.Version

if (-not $osName.Contains("Windows 11")) {
    Write-ColorMessage "Ce script est conçu uniquement pour Windows 11. Système détecté : $osName" "Red"
    exit 1
}
else {
    Write-ColorMessage "Système compatible détecté : $osName" "Green"
}

# Installation du module PSWindowsUpdate si nécessaire
$moduleInstalled = Install-ModuleIfNotExists -ModuleName "PSWindowsUpdate"
if (-not $moduleInstalled) {
    Write-ColorMessage "Impossible de continuer sans le module PSWindowsUpdate." "Red"
    exit 1
}

# Configuration de Windows Update
Write-ColorMessage "`nConfiguration de Windows Update..." "Cyan"
try {
    # Enregistrement du service Windows Update
    Add-WUServiceManager -ServiceID "7971f918-a847-4430-9279-4a52d1efe18d" -Confirm:$false | Out-Null
    Write-ColorMessage "Service Windows Update configuré avec succès." "Green"
}
catch {
    Write-ColorMessage "Erreur lors de la configuration de Windows Update : $_" "Red"
}

# Recherche et installation des mises à jour Windows
Write-ColorMessage "`nRecherche des mises à jour Windows..." "Cyan"
try {
    $availableUpdates = Get-WindowsUpdate -MicrosoftUpdate -NotCategory "Drivers" -ErrorAction Stop
    
    if ($availableUpdates.Count -eq 0) {
        Write-ColorMessage "Aucune mise à jour Windows disponible." "Green"
    }
    else {
        Write-ColorMessage "Nombre de mises à jour Windows disponibles : $($availableUpdates.Count)" "Yellow"
        
        Write-ColorMessage "`nInstallation des mises à jour Windows..." "Cyan"
        $result = Install-WindowsUpdate -MicrosoftUpdate -NotCategory "Drivers" -AcceptAll -IgnoreReboot -ErrorAction Stop
        
        foreach ($update in $result) {
            if ($update.Result -eq "Installed") {
                $updateReport.WindowsUpdates.Installed += "$($update.Title) (KB$($update.KB))"
            }
            else {
                $updateReport.WindowsUpdates.Failed += "$($update.Title) (KB$($update.KB)) - $($update.Result)"
            }
        }
    }
}
catch {
    Write-ColorMessage "Erreur lors de la recherche ou de l'installation des mises à jour Windows : $_" "Red"
}

# Recherche et installation des mises à jour de pilotes
Write-ColorMessage "`nRecherche des mises à jour de pilotes..." "Cyan"
try {
    $availableDrivers = Get-WindowsUpdate -MicrosoftUpdate -Category "Drivers" -ErrorAction Stop
    
    if ($availableDrivers.Count -eq 0) {
        Write-ColorMessage "Aucune mise à jour de pilotes disponible." "Green"
    }
    else {
        Write-ColorMessage "Nombre de mises à jour de pilotes disponibles : $($availableDrivers.Count)" "Yellow"
        
        Write-ColorMessage "`nInstallation des mises à jour de pilotes..." "Cyan"
        $result = Install-WindowsUpdate -MicrosoftUpdate -Category "Drivers" -AcceptAll -IgnoreReboot -ErrorAction Stop
        
        foreach ($driver in $result) {
            if ($driver.Result -eq "Installed") {
                $updateReport.Drivers.Installed += "$($driver.Title)"
            }
            else {
                $updateReport.Drivers.Failed += "$($driver.Title) - $($driver.Result)"
            }
        }
    }
}
catch {
    Write-ColorMessage "Erreur lors de la recherche ou de l'installation des mises à jour de pilotes : $_" "Red"
}

# Vérification si un redémarrage est nécessaire
Write-ColorMessage "`nVérification si un redémarrage est nécessaire..." "Cyan"
$rebootRequired = $false

try {
    # Méthode 1 : Vérification via PSWindowsUpdate
    $rebootStatus = Get-WURebootStatus -Silent
    if ($rebootStatus) {
        $rebootRequired = $true
    }
}
catch {
    Write-ColorMessage "Erreur lors de la vérification du statut de redémarrage via PSWindowsUpdate : $_" "Yellow"
    
    try {
        # Méthode 2 : Vérification via registre
        $regKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
        if (Test-Path $regKey) {
            $rebootRequired = $true
        }
        
        # Méthode 3 : Vérification via CBS
        $pendingRenames = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
        if ($pendingRenames -ne $null) {
            $rebootRequired = $true
        }
    }
    catch {
        Write-ColorMessage "Erreur lors de la vérification alternative du statut de redémarrage : $_" "Yellow"
    }
}

$updateReport.RestartRequired = $rebootRequired

# Affichage du rapport
Show-UpdateReport

# Proposition de redémarrage si nécessaire
if ($updateReport.RestartRequired) {
    $restart = Read-Host "Un redémarrage est nécessaire pour terminer l'installation des mises à jour. Souhaitez-vous redémarrer maintenant ? (O/N)"
    if ($restart -eq "O" -or $restart -eq "o") {
        Write-ColorMessage "Redémarrage du système dans 10 secondes..." "Yellow"
        Start-Sleep -Seconds 10
        Restart-Computer -Force
    }
    else {
        Write-ColorMessage "N'oubliez pas de redémarrer votre système ultérieurement pour finaliser l'installation des mises à jour." "Yellow"
    }
}

<#
SUGGESTION D'AMÉLIORATION :
Pour améliorer ce script, il serait possible d'ajouter :
1. Une option de planification pour exécuter le script régulièrement via une tâche programmée
2. L'envoi d'un rapport par e-mail après l'exécution
3. La gestion des mises à jour pour les applications Microsoft Store
4. Une option pour bloquer certaines mises à jour problématiques via leurs numéros KB
5. L'intégration de la sauvegarde automatique du système avant installation
#> 