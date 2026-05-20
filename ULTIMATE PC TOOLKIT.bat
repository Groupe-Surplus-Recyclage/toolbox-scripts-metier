@echo off
:: Force l'encodage en UTF-8 pour éviter les problèmes d'accents dans la console
chcp 65001 >nul
title BOÎTE À OUTILS PC

:menu
cls
color 0A
echo ===================================================
echo             BOÎTE À OUTILS PC
echo ===================================================
echo 1. Nettoyer les fichiers temp   11. Rapport de batterie
echo 2. Vider le cache DNS           12. Utilisation du disque
echo 3. Infos IP                     13. Ouvrir le Panneau de config
echo 4. Infos Système                14. Ouvrir le Gestionnaire des tâches
echo 5. Vérifier le disque (CHKDSK)  15. Ouvrir les Services
echo 6. Réparer les fichiers (SFC)   16. Ouvrir le Gestionnaire de périph.
echo 7. Redémarrer l'Explorateur     17. Afficher les mots de passe Wi-Fi
echo 8. Réinitialiser le Réseau      18. Redémarrer le PC
echo 9. Afficher les tâches en cours 19. Éteindre le PC
echo 10. Fermer Chrome               20. Quitter
echo ===================================================
echo.
set /p opt="Sélectionnez une option : "

if "%opt%"=="1" goto opt1
if "%opt%"=="2" goto opt2
if "%opt%"=="3" goto opt3
if "%opt%"=="4" goto opt4
if "%opt%"=="5" goto opt5
if "%opt%"=="6" goto opt6
if "%opt%"=="7" goto opt7
if "%opt%"=="8" goto opt8
if "%opt%"=="9" goto opt9
if "%opt%"=="10" goto opt10
if "%opt%"=="11" goto opt11
if "%opt%"=="12" goto opt12
if "%opt%"=="13" goto opt13
if "%opt%"=="14" goto opt14
if "%opt%"=="15" goto opt15
if "%opt%"=="16" octave16
if "%opt%"=="16" goto opt16
if "%opt%"=="17" goto opt17
if "%opt%"=="18" goto opt18
if "%opt%"=="19" goto opt19
if "%opt%"=="20" goto exit
goto menu

:opt1
cls
echo Nettoyage des fichiers temporaires...
del /q /f /s %TEMP%\* 2>nul
del /q /f /s C:\Windows\Temp\* 2>nul
echo Nettoyage terminé.
pause
goto menu

:opt2
cls
echo Vidage du cache DNS...
ipconfig /flushdns
pause
goto menu

:opt3
cls
echo Informations IP actuelles :
ipconfig /all
pause
goto menu

:opt4
cls
echo Récupération des informations système (veuillez patienter)...
systeminfo
pause
goto menu

:opt5
cls
echo Analyse du disque C: (Mode lecture seule)...
chkdsk C:
pause
goto menu

:opt6
cls
echo Analyse et réparation des fichiers système...
sfc /scannow
pause
goto menu

:opt7
cls
echo Redémarrage de l'Explorateur Windows...
taskkill /f /im explorer.exe
start explorer.exe
echo Explorateur Windows redémarré avec succès.
pause
goto menu

:opt8
cls
echo Réinitialisation des protocoles réseau...
netsh int ip reset
netsh winsock reset
echo Réinitialisation terminée. Un redémarrage est fortement conseillé.
pause
goto menu

:opt9
cls
echo Liste des processus et tâches en cours :
tasklist
pause
goto menu

:opt10
cls
echo Fermeture forcée de Google Chrome...
taskkill /f /im chrome.exe 2>nul
echo Chrome a été arrêté.
pause
goto menu

:opt11
cls
echo Génération du rapport de batterie...
powercfg /batteryreport /output "%USERPROFILE%\Desktop\rapport_batterie.html"
echo Le rapport a été enregistré sur votre Bureau sous le nom "rapport_batterie.html".
pause
goto menu

:opt12
cls
echo Analyse de l'espace disque (Lecteur C:) :
dir C:\ | find "fichiers"
dir C:\ | find "rep(s)"
pause
goto menu

:opt13
cls
echo Ouverture du Panneau de Configuration...
control
goto menu

:opt14
cls
echo Ouverture du Gestionnaire des tâches...
start taskmgr.exe
goto menu

:opt15
cls
echo Ouverture des Services Windows...
start services.msc
goto menu

:opt16
cls
echo Ouverture du Gestionnaire de périphériques...
start devmgmt.msc
goto menu

:opt17
cls
echo Extraction des mots de passe Wi-Fi enregistrés :
echo.
for /f "tokens=2 delims=:" %%i in ('netsh wlan show profiles ^| findstr "Profil"') do (
    set "ssid=%%i"
    setlocal enabledelayedexpansion
    set "ssid=!ssid:~1!"
    echo Réseau : !ssid!
    netsh wlan show profile name="!ssid!" key=clear | findstr "Contenu"
    echo ----------------------------------------
    endlocal
)
pause
goto menu

:opt18
cls
echo Le PC va redémarrer dans 10 secondes. Fermez vos fichiers en cours !
shutdown /r /t 10
pause
goto menu

:opt19
cls
echo Le PC va s'éteindre dans 10 secondes. Fermez vos fichiers en cours !
shutdown /s /t 10
pause
goto menu

:exit
exit