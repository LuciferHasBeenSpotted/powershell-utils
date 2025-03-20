$UEFI_CHECK = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State" -Name "UEFISecureBootEnabled" -ErrorAction SilentlyContinue
 
$SN_SCREENS = @()
$EXCEL = New-Object -ComObject Excel.Application
$EXCEL.Visible = $true
 
$FIRST_NAME = $env:USERNAME.Split('.')[0]
$LAST_NAME = $env:USERNAME.Split('.')[1]
 
$TENANT_ID = "..."
$CLIENT_ID = "..."
$CLIENT_SECRET = "..."
 
$TOKEN_URL = "https://login.microsoftonline.com/$TENANT_ID/oauth2/v2.0/token"
 
$BODY = @{
    client_id = $CLIENT_ID
    client_secret = $CLIENT_SECRET
    grant_type = "client_credentials"
    scope = "https://graph.microsoft.com/.default"
}
 
$RES_TOKEN = Invoke-RestMethod -Uri $TOKEN_URL -Method Post -ContentType "application/x-www-form-urlencoded" -Body $BODY
$TOKEN = $RES_TOKEN.access_token
 
$SITE_ID = "..."
$FILE_PATH = "..."

$LOCAL_PATH = "$([System.IO.Path]::GetTempPath())info-support.xlsx"

# Telecharger le fichier excel depuis le SP
$DOWNLOAD_URI = "https://graph.microsoft.com/beta/sites/$SITE_ID/drive/root:/" + $FILE_PATH + ":/content"
Invoke-RestMethod -Uri $DOWNLOAD_URI -Headers @{Authorization = "Bearer $TOKEN"} -Method Get -OutFile $LOCAL_PATH

function Save-Excel {
    param ([System.Object]$Workbook,
        [System.Object]$Excel,
        [System.Object]$Worksheet)
 
    $Workbook.Save()
    $Workbook.Close()
    $Excel.Quit()
 
    #Modifier le fichier excel du SP par la version modifiée dans le %TEMP%
    Invoke-RestMethod -Uri $DOWNLOAD_URI -Method PUT -Headers @{Authorization = "Bearer $TOKEN"} -InFile $LOCAL_PATH -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" | Out-Null
 
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
}

function Write-Screens {
    param([bool]$FirstTime,
    [int]$Row,
    [System.Object]$Worksheet)
    
    if($FirstTime) {
        $Worksheet.Cells.Item($Row, 1).Value = $FIRST_NAME
        $Worksheet.Cells.Item($Row, 2).Value = $LAST_NAME
        $Worksheet.Cells.Item($Row, 4).Value = $env:COMPUTERNAME
        $Worksheet.Cells.Item($Row, 5).Value = (Get-WmiObject -class win32_bios).SerialNumber
        $Worksheet.Cells.Item($Row, 6).Value = @("LEGACY BIOS", "UEFI")[$UEFI_CHECK.UEFISecureBootEnabled -eq 1]
    }
    
    $Worksheet.Cells.Item($Row, 3).Value = Get-Date -UFormat "%Y-%m-%d %H:%M:%S"
    #Obtenir la derniere cellule de la ligne 
    $lastCell = $Worksheet.Cells($Row, $Worksheet.Columns.Count).End(-4159).Column
    
    # recuperer les sn deja dans le tableau
    $range = "G$Row"+":" + [char](64 + $lastCell + 1) + "$Row"
    
    # verifier si les ecrans branches ne sont pas deja dans le fichier excel
    $new_screens = ($Worksheet.Range($range).Value2 + $SN_SCREENS) | Select-Object -Unique

    for($i = 0; $i -lt $new_screens.Length; $i++) {
        $Worksheet.Cells.Item($Row, 7 + $i).Value = $new_screens[$i]
    }
}

$WORKBOOK = $EXCEL.Workbooks.Open($LOCAL_PATH)
$WORKSHEET = $WORKBOOK.Sheets.Item(1)
#Obtenir la derniere ligne vide du tableau
$LAST_ROW = $WORKSHEET.Cells($WORKSHEET.Rows.Count, 1).End(3).Row + 1

# incementer le tableau des ecrans sans l ecran du pc portable
forEach($screen in gwmi WmiMonitorID -Namespace root\wmi) {
    if($null -eq $screen.UserFriendlyName) { continue }
    $SN_SCREENS += ($screen.SerialNumberID | ForEach-Object {[char]$_}) -join ""
}

# Boucle sur chaque ligne pour chercher l'utilisateur
for($index_row = 2; $index_row -lt $LAST_ROW; $index_row++) {
    $prenom = $WORKSHEET.Cells.Item($index_row, 1).Text
    $nom = $WORKSHEET.Cells.Item($index_row, 2).Text
 
    if ($prenom -eq $FIRST_NAME -and $nom -eq $LAST_NAME) {
        Write-Screens -FirstTime $false -Worksheet $WORKSHEET -Row $index_row
        Save-Excel -Workbook $WORKBOOK -Excel $EXCEL -Worksheet $WORKSHEET

        # arret du script ici si user dans le fichier excel
        exit
    }
}

# si user pas dans le fichier
Write-Screens -FirstTime $true -Worksheet $WORKSHEET -Row $LAST_ROW
Save-Excel -Workbook $WORKBOOK -Excel $EXCEL -Worksheet $WORKSHEET
