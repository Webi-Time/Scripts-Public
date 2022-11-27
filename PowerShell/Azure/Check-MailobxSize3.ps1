<#
    .Synopsis
       Permet de se connecter au tenant via Microsoft Graph et de verifier la taille des boites au lettre et d'alerter un administrateurs.

    .DESCRIPTION
       Ce script powershell fait appel à l'API Microsoft Graph. Il permet de se connecter puis recuperer la taille de toutes les boites Mail
       et de verifiéque celle ci ne depasse pas un seuil, dans le cas contreire,un mail est envoyer pour alerter un Adminsitrateur avec la liste des
       utilisateurs qui depasse le seuil.
       Il se deconnecter ensuite du tenant.

    .EXAMPLE
       .\Check-MailboxSize.ps1

    .EXAMPLE
       .\Check-MailboxSize.ps1  -VerboseLvl 1

    .INPUTS
       Un fichier JSON est necessaire pour fonctionner et se connecter au tenant. 
       Ce fichier contient :
        - clientId          : Application (client) ID
        - clientCertificate : Le thundprint du certificat
        - tenantId          : Directory (tenant) ID
        - organisationName  : domaine.onmicrosoft.com

    .INPUTS Optionnal        
        - VerboseLvl  : 0 par defaut : Verbose [0 = Aucun ; 1 = Brief ; 2 = Tout], afficher les informtions dans la console

    .OUTPUTS
       

    .NOTES
       N.A

    .FUNCTIONALITY
       N.A
#>


[cmdletbinding()]
Param (
    [ValidateSet(0,1,2)][System.Byte]$VerboseLvl = 0
)


#region Variables
    $Global:VerboseLvl = $VerboseLvl
    $Global:sTimestamp = Get-Date -F "yyyyMMdd_HHmmss"
    $Global:sRunPath = Get-Location
    $Global:sRunUser = $env:UserName
    $Global:sRunUserDomain = $env:UserDomain
    $Global:logFileDir = "$sRunPath\Logs"
    $Global:LogSource = "PowerShell-Script-Check_MailboxSize"
#endregion Variables


#region functions

    Import-Module .\GeneriqueFunction.ps1 -Force

#endregion functions


#region Event message

$Message_Event_Error_GetStockage = # Event 2573
"   Utilisateur `t: $sRunUserDomain\$sRunUser
    Description `t: Impossible de récuperer le stockage utilisateurs
    Date `t`t`t: $(Get-Date)
    Information complementaire :`n`r 
" 
$Message_Event_Error_SupprCsv = # Event 2574
"   Utilisateur `t: $sRunUserDomain\$sRunUser
    Description `t: Impossible de supprimer le CSV
    Date `t`t`t: $(Get-Date)
    Information complementaire :`n`r
" 

#endregion Event message


#region main

    Clear-Host
    Log "Debut du script [$($MyInvocation.MyCommand.Name)] " $logFileDir 0 Magenta

    $oldErrorActionPreference = $ErrorActionPreference;
    $ErrorActionPreference = "Stop"
    

    ## Import du module Microsoft.Graph
    Import-MsGraph
    
    
    # Connexion au tenant
    Connect-MsGraphTenant

    # Verifie la taille des boites au lettre

        $seuilMail = "1MB" 

        $Periode = "D30"
        $CsvPath = $PSScriptRoot
        $CsvFile = "ReportMailboxUsageDetail-$Periode-$(Get-Date -Format 'yyyy_MM_dd-hh_mm_ss').csv"
        
        
        try { 
            Log "Récuperation du stockage utilisé par les utilisateurs dans un fichier temporaire " $logFileDir 1 Cyan
            Get-MgReportMailboxUsageDetail -Period $Periode -OutFile "$($CsvPath)\$CsvFile" -ErrorAction Stop
            $Csv = Import-Csv -Path "$($CsvPath)\$CsvFile" -Delimiter ','
            $CsvTrie = $Csv | Where-Object {[int64]($_.'Storage Used (Byte)') -gt $seuilMail}
        } catch { 
            Log "Erreur - Impossible de récuperer le stockage utilisé par les utilisateurs `r`n`t Message d'erreur : $($_)" $logFileDir 0 Red
            New-ComgestEvent -Message "$Message_Event_Error_GetStockage `t Fichier : [$($CsvPath)\$CsvFile]`r`n`t Message  d'erreur : $($_)" -EventID 2573 -EventInstance Error -Source $LogSource
            Disconnect-MsGraphTenant
            exit 1  
        } 

        try { 
            Log "Suppression du fichier temporaire" $logFileDir 1 Cyan
            Remove-Item -Path "$($CsvPath)\$CsvFile" -Confirm:$false
        } catch { 
            Log "Erreur - Impossible de supprimer le fichier temporaire `r`n`t Message d'erreur : $($_)" $logFileDir 0 Red
            New-ComgestEvent -Message "$Message_Event_Error_SupprCsv `t Fichier : [$($CsvPath)\$CsvFile]`r`n`t Message d'erreur`t: $($_)" -EventID 2574 -EventInstance Error -Source $LogSource  
        } 


        if ($CsvTrie){
            try { 
                Log "Des boites mail qui dépasse $seuilMail ont été trouvés.. " $logFileDir 0 Yellow
                $b = $CsvTrie | Sort-Object TotalItemSize -Descending  | Select-Object DisplayName, TotalItemSize | ConvertTo-Html       
                Send-Mailmessage -smtpServer smtp.bur.comgest-sa.com -from IT-Notify@comgest.com -to hd@comgest.com -subject "Mailbox size more than 90GB" -body "$b" -priority High -ErrorAction Stop -BodyAsHtml
            } catch { 
                Log "Erreur - Impossible d'envoyer le mail `r`n`t Message d'erreur : $($_)" $logFileDir 0 Red
                New-ComgestEvent -Message "$Message_Event_Error_SendMail `t Message d'erreur : $($_)" -EventID 2572 -EventInstance Error -Source $LogSource
                Disconnect-MsGraphTenant
                exit 1   
            } 
        }else{
            Log "Aucune boite mail ne dépasse $($seuilMail/1GB) " $logFileDir 0 Green
        }
        
    # Connexion au tenant
    Disconnect-MsGraphTenant
    Exit 0

#endregion Main