#Importations
$cache_path = "C:\Users\$env:USERNAME\AppData\Local\Google\Chrome\User Data\Default\Cache\Cache_Data"
$chrome_path = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe";

#Recuperer le processus Chrome et forcer sa fermeture
Get-Process -Name chrome | Stop-Process -Force

#Recuperer les fichiers du dossier de cache | Filtre sur les noms != data* & index* | Forcer la suppression
Get-ChildItem -Path $cache_path | Where-Object {$_.Name -notmatch '^data' -and $_.Name -notmatch '^index'} | Remove-Item -Force

#Relancer le processus de chrome
Start-Process -FilePath $chrome_path
