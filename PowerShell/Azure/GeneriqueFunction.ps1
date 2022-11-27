#region Events Generale

    ##
    ## SUCCESS
    ##
    $Message_Event_Success_SignIn = # Event 2560
    "   Utilisateur `t: $sRunUserDomain\$sRunUser
        Description `t: Connexion au tenant effectué par le script 
        Date`t: $(Get-Date)
    " 
    $Message_Event_Success_SignOut = # Event 2561
    "   Utilisateur `t: $sRunUserDomain\$sRunUser
        Description `t: Deconnexion au tenant effectué par le script 
        Date`t: $(Get-Date)
    "  
    ##
    ## ERROR
    ##
    $Message_Event_Error_SignIn = # Event 2570
    "   Utilisateur `t: $sRunUserDomain\$sRunUser
        Description `t: Impossible de se connecter 
        Date`t: $(Get-Date)
        Information complementaire sur l'erreur :`n`r
    " 
    $Message_Event_Error_SignOut = # Event 2571
    "   Utilisateur `t: $sRunUserDomain\$sRunUser
        Description `t: Impossible de se déconnecter 
        Date`t: $(Get-Date)
        Information complementaire sur l'erreur :`n`r
    "  
    $Message_Event_Error_SendMail = # Event 2572
    "   Utilisateur `t: $sRunUserDomain\$sRunUser
        Description `t: Impossible d'envoyer le mail
        Date`t: $(Get-Date)
        Information complementaire :`n`r
    " 
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
    $organisation = $settings.organisationName

    ## Déchargement du module 365 et déconnexion 
    try
    { 
        Log "Tentative de connexion au tenant " $logFileDir 2 Cyan
        Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $certificate -ForceRefresh -ErrorAction Stop | Out-Null   
        New-ComgestEvent -Message $Message_Event_Success_SignIn -EventID 2560 -EventInstance Information -Source $LogSource
        Log "Connexion au tenant effectuée" $Global:logFileDir 1 Green
    }
    catch 
    {
        Log "Erreur - Impossible de se connecter au tenant `r`n`t`t  Message d'erreur : $($_)" $logFileDir 0 Red
        New-ComgestEvent -Message "$Message_Event_Error_SignIn `t Message d'erreur : $($_)" -EventID 2570 -EventInstance Error -Source $LogSource
        exit 1
    }
}

function Disconnect-MsGraphTenant {
	## Déchargement du module 365 et déconnexion 
    try{
        Disconnect-MgGraph -ErrorAction Stop  | Out-Null     
        Log "Tentative de deconnexion du tenant " $logFileDir 2 Cyan
        New-ComgestEvent -Message $Message_Event_Success_SignOut -EventID 2561 -EventInstance Information -Source $LogSource
        Log "Deconnexion du tenant effectué" $logFileDir 1 Green
    } catch {
        Log "Erreur - Impossible de se deconnecter au tenant`r`n`t`t  Message d'erreur : $($_)" $logFileDir 0 Red            
        New-ComgestEvent -Message "$Message_Event_Error_SignOut `t Message d'erreur : $($_)" -EventID 2571 -EventInstance Error -Source $LogSource
        exit 1
    }
}

function New-ComgestEvent {
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

        if (-not (Test-Path -Path $sLogDir)) {
            New-Item -ItemType Directory -Path $logFileDir
        }

	    $sLogFile = $sLogDir + "\" + "Log_" + $sTimeStamp + ".log"
	    $sLineTimeStamp = Get-Date -f "dd/MM/yyyy HH:mm:ss"
	    $sInput | % {
		    $sLine = $sLineTimeStamp + " - " + $_		    
		    $sLine | Out-File $sLogFile -Append:$true -Force
	    }
        if ($lvl -le $VerboseLvl ){Write-Host $sLine -ForegroundColor $color}

    }

function Import-MsGraph ([ValidateSet("CurrentUser", "AllUsers")][System.String]$scope = "AllUsers",[boolean]$RefreshInstallModule = $false){
    Log "Verification du module MSGraph" $logFileDir 1 Cyan
    $module = (Get-Module -ListAvailable | where {$_.Name -like "Microsoft.Graph*"})
    if (([string]::IsNullOrEmpty($module) -or $module.count -le 2) -or $RefreshInstallModule) # Si pas installé
    {
        Log "Installation du module MSGraph" $logFileDir 1 Yellow
        Import-Module PowerShellGet -ErrorAction SilentlyContinue
        Install-Module Microsoft.Graph -Scope $scope -Confirm:$false -Force -ErrorAction SilentlyContinue | Out-Null
        sleep 5
        try{
            Log "Chargement du module MSGraph" $logFileDir 2 Cyan
            Import-Module -Name Microsoft.Graph -Force | Out-Null
        }
        catch {
            Log "Erreur - Impossible de charger le module completement`r`n`t`t  Message d'erreur : $($_)" $logFileDir 0 Red 
            Log "Le module est chargée en partie. Voici la liste des modules chargés :" $logFileDir 1 Yellow
            Log $($(Get-Command -Module Microsoft.Graph* | select Source -Unique) | Out-String) $logFileDir 2 Yellow
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
            Log $($(Get-Command -Module Microsoft.Graph* | select Source -Unique) | Out-String) $logFileDir 2 Yellow 
        }
    }
    else # si installer et importé
    {
        Log "Le module MSGraph est deja chargé" $logFileDir 2 Cyan
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