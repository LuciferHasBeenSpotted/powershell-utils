#lien vers le setup.exe
$url = "..."

#/temp/setup.exe
$out_path = "$([System.IO.Path]::GetTempPath())setup.exe"

#Verifier si l'app n'est pas déjà installé 
$install = Test-Path "..."

if(!$install) {
    Write-Host "Telechargement..."

    Invoke-WebRequest -Uri $url -OutFile $out_path #Recupere le Setup

    Start-Process -FilePath $out_path -Wait #Exectute le setup

    Start-Sleep -Seconds 5

    Remove-Item $out_path -Force #Desinstaller le setup du /temp

    Write-Host "Telechargement fini, lancement"
}
