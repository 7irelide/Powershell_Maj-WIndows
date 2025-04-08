@echo off
echo ===============================================
echo    LANCEMENT DES MISES A JOUR WINDOWS 11
echo ===============================================
echo.

:: Vérification de l'exécution en mode administrateur
NET SESSION >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Ce script doit etre execute en tant qu'administrateur.
    echo Demande d'elevation des privileges...
    
    :: Auto-élévation des privilèges
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo Execution du script de mise a jour Windows 11...
echo.

:: Détermination du chemin du script PowerShell
set "SCRIPT_PATH=%~dp0Update-Windows11.ps1"

:: Vérification de l'existence du script PowerShell
if not exist "%SCRIPT_PATH%" (
    echo ERREUR: Le script Update-Windows11.ps1 est introuvable.
    echo Assurez-vous qu'il se trouve dans le meme repertoire que ce fichier batch.
    echo Chemin recherche: %SCRIPT_PATH%
    echo.
    pause
    exit /b 1
)

:: Définition de la politique d'exécution et lancement du script PowerShell
powershell -ExecutionPolicy Bypass -Command "& '%SCRIPT_PATH%'"

echo.
echo Fin de l'execution du script.
echo.
pause 