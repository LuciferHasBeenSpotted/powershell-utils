Import-Module ActiveDirectory
 
$csvPath = ""
$lignes = Import-Csv -Path $csvPath
$nomGroupe = ""
 
foreach($ligne in $lignes.users) {
    $ligne = $ligne -split " "
   

    if($ligne.Count -gt 2) { Write-Host "nom composé: $ligne" }
    else {
        $prenom = $ligne[0]
        $nom = $ligne[1]

        $utilisateur = Get-ADUser -Filter "SamAccountName -eq '$prenom.$nom'"
        if(-not $utilisateur) { Write-Host "Utilisateur introuvable: $ligne" }
        else {Add-ADGroupMember -Identity $nomGroupe -Members $utilisateur }
    }
}
