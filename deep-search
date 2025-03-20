$CHEMIN_ARCHIVE = Read-Host "Entrez le chemin vers l'archive"
$NOM_DOSSIER = Read-Host "Entrez le nom du fichier ou dossier recherche"
$EXTRACION = Read-Host "Entrez la destination de l'extraction"

function validate-element {
    [CmdletBinding()] Param([System.IO.FileSystemInfo] $nom)

    # si l'élement donné est un dossier
    if($nom.PSIsContainer) {
        if($nom -like "*$NOM_DOSSIER*") {
            Copy-Item -Path $nom.FullName -Destination $EXTRACION -Force -Recurse
            Write-Host $nom.FullName
        } else {
            foreach($el in Get-ChildItem -path $nom.FullName) {
                $fn = $nom.FullName
                Write-Host "J'analyse le répertoire: $fn"
                validate-element $el
            }
        }
    } else {
        # si l'élement donné est un fichier
        if($nom -like "*$NOM_DOSSIER*") {
            Copy-Item -Path $nom.FullName -Destination $EXTRACION -Force
            Write-Host $nom.FullName
        }
    }
}

foreach($element in Get-ChildItem -path $CHEMIN_ARCHIVE) {
    validate-element $element 
}
