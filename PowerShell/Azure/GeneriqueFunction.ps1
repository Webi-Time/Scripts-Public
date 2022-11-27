#region Events Generale

    ##
    ## SUCCESS
    ##
    $Event_Success_SignIn = @{ 
        message = "
            Utilisateur`t : $sRunUserDomain\$sRunUser
            Script`t : $sRunScript
            Description : Connexion au tenant effectuée par le script
            Date`t : $(Get-Date)
        "
        id = 2560
        EventType = 'Information'
        LogSource = $LogSource
    }
    $Event_Success_SignOut = @{ 
        message = "
            Utilisateur`t : $sRunUserDomain\$sRunUser
            Script`t : $sRunScript
            Description : Deconnexion du tenant effectuée par le script
            Date`t : $(Get-Date)
        "
        id = 2561
        EventType = 'Information'
        LogSource = $LogSource
    }

    ##
    ## ERROR
    ##
    $Event_Error_SignIn = @{ 
        message = "
            Utilisateur`t : $sRunUserDomain\$sRunUser
            Script`t : $sRunScript
            Description : Impossible de se connecter au tenant 
            Date`t : $(Get-Date)
            Information complementaire sur l'erreur :`n`r
        "
        id = 2570
        EventType = 'Error'
        LogSource = $LogSource
    }
    $Event_Error_SignOut = @{ 
        message = "
            Utilisateur`t : $sRunUserDomain\$sRunUser
            Script`t : $sRunScript
            Description : Impossible de se déconnecter du tenant 
            Date`t : $(Get-Date)
            Information complementaire sur l'erreur :`n`r
        "
        id = 2571
        EventType = 'Error'
        LogSource = $LogSource
    }
    $Event_Error_SendMail = @{ 
        message = "
            Utilisateur`t : $sRunUserDomain\$sRunUser
            Script`t : $sRunScript
            Description : Impossible d'envoyer le mail
            Date`t : $(Get-Date)
            Information complementaire sur l'erreur :`n`r
        "
        id = 2572
        EventType = 'Error'
        LogSource = $LogSource
    }

#endregion Events Generale


#region Function generique

function Connect-MsGraphTenant {	
    # Se positionne dans le meme dossier que le script
    Set-Location $PSScriptRoot
  
    # Chargement des parametres depuis le fichier JSON ##
    Log "Chargement des parametres ..." $Global:logFileDir 2 Cyan
    $settings = Get-Content '.\settings.json' -ErrorAction Stop | Out-String | ConvertFrom-Json   
    #$settings = Get-Content '.\SettingsDynamics.json' -ErrorAction Stop | Out-String | ConvertFrom-Json   
    $clientId = $settings.clientId
    $tenantId = $settings.tenantId
    $certificate = $settings.clientCertificate
    #$organisation = $settings.organisationName

    ## Déchargement du module 365 et déconnexion 
    try
    { 
        Log "Tentative de connexion au tenant " $logFileDir 2 Cyan
        Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $certificate -ForceRefresh -ErrorAction Stop | Out-Null
        New-CustomEvent -Message $Event_Success_SignIn.message -EventID $Event_Success_SignIn.id -EventInstance $Event_Success_SignIn.EventType -Source $Event_Success_SignIn.LogSource
        Log "Connexion au tenant effectuée" $Global:logFileDir 1 Green
    }
    catch 
    {
        Log "Erreur - Impossible de se connecter au tenant `r`n`t`t  Message d'erreur : $($_)" $logFileDir 0 Red
        New-CustomEvent -Message "$($Event_Error_SignIn.message) `t Message d'erreur : $($_)" -EventID $Event_Error_SignIn.id -EventInstance $Event_Error_SignIn.EventType -Source $Event_Error_SignIn.LogSource
        exit 1
    }
}

function Disconnect-MsGraphTenant {
	## Déchargement du module 365 et déconnexion 
    try{
        Disconnect-MgGraph -ErrorAction Stop  | Out-Null
        Log "Tentative de deconnexion du tenant " $logFileDir 2 Cyan
        New-CustomEvent -Message $Event_Success_SignOut.message -EventID $Event_Success_SignOut.id -EventInstance $Event_Success_SignOut.EventType -Source $Event_Success_SignOut.LogSource
        Log "Deconnexion du tenant effectué" $logFileDir 1 Green
    } catch {
        Log "Erreur - Impossible de se deconnecter au tenant`r`n`t`t  Message d'erreur : $($_)" $logFileDir 0 Red
        New-CustomEvent -Message "$($Event_Error_SignOut.message) `t Message d'erreur : $($_)" -EventID $Event_Error_SignOut.id -EventInstance $Event_Error_SignOut.EventType -Source $Event_Error_SignOut.LogSource
        exit 1
    }
}

function New-CustomEvent {
    param
    ( 
        [Parameter(Position=0, Mandatory = $false)] $EventLog  = "Application",
        [Parameter(Position=1, Mandatory = $true)] $Source ,
        [Parameter(Position=2, Mandatory = $true)] $EventID   = "2560",
        [Parameter(Position=3, Mandatory = $true) ] $Message,
        [Parameter(Position=4, Mandatory = $false)][ValidateSet("Information","Warning","Error")] $EventInstance = 'Warning'
    )      
    If ([System.Diagnostics.EventLog]::SourceExists($Source) -eq $false) {[System.Diagnostics.EventLog]::CreateEventSource($Source, $EventLog)}
  
    Switch ($EventInstance){
        {$_ -match 'Error'}       {$id = New-Object System.Diagnostics.EventInstance($EventID,1,1)} #ERROR EVENT
        {$_ -match 'Warning'}     {$id = New-Object System.Diagnostics.EventInstance($EventID,1,2)} #WARNING EVENT
        {$_ -match 'Information'} {$id = New-Object System.Diagnostics.EventInstance($EventID,1)}   #INFORMATION EVENT
    } 
    $Object = New-Object System.Diagnostics.EventLog;
    $Object.Log       = $EventLog;
    $Object.Source    = $Source;
 
    $Object.WriteEvent($id, @($Message))
        
}

Function Log {
	    Param ( $sInput, $sLogDir, $lvl ,[string]$color = "Cyan")
	    $sLogFile = $sLogDir + "\" + "Log_$($sRunScript)_" + $sTimeStamp + ".log"
	    $sLineTimeStamp = Get-Date -f "dd/MM/yyyy HH:mm:ss"
	    $sInput | ForEach-Object {
		    $sLine = $sLineTimeStamp + " - " + $_		    
		    $sLine | Out-File $sLogFile -Append:$true -Force
	    }
        if ($lvl -le $VerboseLvl ){Write-Host $sLine -ForegroundColor $color}

    }

function Import-MsGraph ([ValidateSet("CurrentUser", "AllUsers")][System.String]$scope = "AllUsers",[boolean]$RefreshInstallModule = $false){
    Log "Verification du module MSGraph" $logFileDir 1 Cyan
    $module = (Get-Module -ListAvailable | Where-Object {$_.Name -like "Microsoft.Graph*"})
    if (([string]::IsNullOrEmpty($module) -or $module.count -le 2) -or $RefreshInstallModule) # Si pas installé
    {
        Log "Installation du module MSGraph" $logFileDir 1 Yellow
        Import-Module PowerShellGet -ErrorAction SilentlyContinue
        Install-Module Microsoft.Graph -Scope $scope -Confirm:$false -Force -ErrorAction SilentlyContinue | Out-Null
        Start-Sleep 5
        try{
            Log "Chargement du module MSGraph" $logFileDir 2 Cyan
            Import-Module -Name Microsoft.Graph -Force | Out-Null
        }
        catch {
            Log "Erreur - Impossible de charger le module completement`r`n`t`t  Message d'erreur : $($_)" $logFileDir 0 Red 
            Log "Le module est chargée en partie. Voici la liste des modules chargés :" $logFileDir 1 Yellow
            Log $($(Get-Command -Module Microsoft.Graph* | Select-Object Source -Unique) | Out-String) $logFileDir 2 Yellow
        }
    }
    elseif ([string]::IsNullOrEmpty((Get-Module | Where-Object {($_.Name -eq "Microsoft.Graph.Users") -or ($_.Name -eq "Microsoft.Graph.Authentication")})))  # Si installé mais pas importé
    {
        try{
            Log "Chargement du module MSGraph" $logFileDir 2 Cyan
            Import-Module -Name Microsoft.Graph -Force | Out-Null
        }
        catch {
            Log "Erreur - Impossible de charger le module completement`r`n`t`t  Message d'erreur : $($_)" $logFileDir 0 Red 
            Log "Le module est chargée en partie. Voici la liste des modules chargés :" $logFileDir 1 Yellow
            Log $($(Get-Command -Module Microsoft.Graph* | Select-Object Source -Unique) | Out-String) $logFileDir 2 Yellow 
        }
    }
    else # si installer et importé
    {
        Log "Le module MSGraph est deja chargé" $logFileDir 2 Cyan
    }
}

function SendMail {
    param (
        [string]$ToMail,
        [string]$MailSubject,
        [string]$MailBody
    )
    Import-Module Microsoft.Graph.Users.Actions
    $params = @{
        Message = @{
            Subject = "$MailSubject"
            Body = @{
                ContentType = "HTML"
                Content = "$MailBody"
            }
            ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = "$ToMail"
                    }
                }
            )
            Importance = "High"
        }
    }
        
    try {       
        Send-MgUserMail -UserId jmarc@webi-time.fr -BodyParameter $params -ErrorAction Stop
    } catch { 
        Log "Erreur - Impossible d'envoyer le mail `r`n`t Message d'erreur : $($_)" $logFileDir 0 Red
        New-CustomEvent -Message "$($Event_Error_SendMail.message) `t Message d'erreur : $($_)" -EventID $Event_Error_SendMail.id -EventInstance $Event_Error_SendMail.EventType -Source $Event_Error_SendMail.LogSource
        Disconnect-MsGraphTenant
        exit 1   
    } 
}

#endregion Function generique


<#

function testMsGraph (){
Set-Location "D:\OneDrive\OneDrive - Magellan Partners\Bureau\Boulot\Comgest\Script-Etape3-MsGraph-Certfifcat"
Measure-Command { .\Start-ConnectTenant3.ps1 2} |select TotalSeconds
pause
Measure-Command { .\Start-ConnectTenant3.ps1 -deconnexion 2} |select TotalSeconds
pause
Measure-Command { .\Check-LastSynchronisation3.ps1 2} |select TotalSeconds
pause
Measure-Command { .\Check-MailobxSize3.ps1 2} |select TotalSeconds
pause
Measure-Command { .\Set-O365CalendarMultiPermission3.ps1 2} |select TotalSeconds
pause
}


function testMsGraph (){
Set-Location "D:\OneDrive\OneDrive - Magellan Partners\Bureau\Boulot\Comgest\Script-Etape3-MsGraph-Certfifcat"
.\Start-ConnectTenant3.ps1 2
pause
.\Start-ConnectTenant3.ps1 -deconnexion 2
pause
.\Check-LastSynchronisation3.ps1 2
pause
.\Check-MailobxSize3.ps1 2
pause
.\Set-O365CalendarMultiPermission3.ps1 2
pause
}


#>