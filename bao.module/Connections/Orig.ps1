
function Connect-SharepointCSOM
{
    Param
    (
        [string]$Room = 'Room-0961',

        [switch]$PassThruSucceeded
    )

    $succeeded = $true

    if ( !$AccorConfig.SharePointClientContext )
    {

        try
        {

            $spCSOMpath = "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\15\ISAPI"

            $AddNecessaryType = {
                Param
                (
                    [Parameter(ValueFromPipeline=$true)]
                    [System.IO.FileInfo]$Path
                )
                Begin
                {
                    $modules = [appdomain]::CurrentDomain.GetAssemblies() |
                        % { $_.GetLoadedModules() } |
                        group Name -AsHashTable
                }
                Process
                {
    
                    if ( ! $modules.ContainsKey( $Path.Name ) )
                    {
                        Add-Type -Path $Path.FullName
                    }
                }

            }
            
            $paths = @()
            $paths += "$spCSOMpath\Microsoft.SharePoint.Client.dll"
            $paths += "$spCSOMpath\Microsoft.SharePoint.Client.Runtime.dll"
            $paths += "$spCSOMpath\Microsoft.SharePoint.Client.Taxonomy.dll"

            $paths | &$AddNecessaryType

            if ($spCred = Get-BaoCredential -Service Room0961)
            {
                $spUser = $spCred.UserName
                $spSecPass = $spCred.Password
                $spUrl = "https://accor.sharepoint.com/sites/Communities-v0/$Room"

                $AccorConfig.SharePointClientContext = New-Object Microsoft.SharePoint.Client.ClientContext($spUrl)

                $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($spUser, $spSecPass)
    
                $AccorConfig.SharePointClientContext.Credentials = $credentials

            }
            else { $succeeded = $false }

        }
        catch
        {
            $succeeded = $false
        }
    }
    if ( $PassThruSucceeded ) { $succeeded }

}

function Connect-SharepointREST
{
    Param
    (
        [string]$Room = 'Room-0961',
        [switch]$PassThruSucceeded
    )

    $succeeded = $false

    if (

        ($AccorConfig.FormDigestExpireAt -is [DateTime]) `        -and
        ((Get-Date).AddMinutes(5) -lt $AccorConfig.FormDigestExpireAt)
    
    ) {

        $succeeded = $true
    
    } elseif (

        $rCred = Get-BaoCredential -Service $Room.Replace('-',[string]::Empty) -ErrorAction SilentlyContinue
    
    ){

        $AccorConfig.SharepointApiUrl = "https://accor.sharepoint.com/sites/Communities-v0/$Room/_api"
        
        $resp = Invoke-SPORestMethod -Url "$($AccorConfig.SharepointApiUrl)/contextinfo" -Method Post
        
        $fdv = $AccorConfig.FormDigestValue = $resp.GetContextWebInformation.FormDigestValue 
        $fdto = $resp.GetContextWebInformation.FormDigestTimeoutSeconds
        
        $AccorConfig.FormDigestExpireAt = ([DateTime]($fdv -split ',')[-1]).AddSeconds($fdto)
        
        $succeeded = $true
    }
    if ( $PassThruSucceeded ) { $succeeded }

}

function _GetCasExchangeServers
{
    Param (
        [ValidateSet('EU', 'AA', 'AU', 'SA', $null)]
        [string]$ContinentCode
    )

    $codeH = @{
        EU = 'EVY01'
        AA = 'SIN02'
        AU = 'SYD02'
        SA = 'SAO01'
    }


    if ($ContinentCode) {

        $f = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
        $fContextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Forest
        $fContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext($fContextType, $f)
        $localSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::FindByName($fContext, $codeH.$ContinentCode)

    } else {

        $localSite=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()

    }


	$configNC=(Get-ADRootDSE).configurationNamingContext

    $localSiteDN = $localSite.GetDirectoryEntry().DistinguishedName
    $adjacentSitesDN = $localSite.AdjacentSites | %{$_.GetDirectoryEntry().DistinguishedName}


    $Filter = "(&(objectClass=msExchExchangeServer)(versionNumber>=1937801568))"
    $exchangeCasObjects = Get-ADObject -LDAPFilter $Filter -SearchBase $configNC -SearchScope Subtree -Properties msExchServerSite,msexchcurrentserverroles,networkaddress |
        ?{ $_.msexchcurrentserverroles -band 4 }


    $exchangeCasServers = $exchangeCasObjects | ?{$_.msExchServerSite -eq $localSiteDN}

    if(!$exchangeCasServers)
    {
        $exchangeCasServers = $exchangeCasObjects | ?{$adjacentSitesDN -contains $_.msExchServerSite}
    }

    if(!$exchangeCasServers)
    {
        $exchangeCasServers = $exchangeCasObjects 
    }


    $exchangeCasServers |
    %{
        $_.networkaddress |
        ? { $_ -like 'ncacn_ip_tcp:*' } |
        % { ($_ -split ':')[1] }
    }
}


function Connect-Exchange
{
Param (
    [ValidateSet('Premises', 'PremisesEU', 'PremisesAA', 'PremisesAU', 'PremisesSA', 'O365')]
    [string]$Service,

    [switch]$ForceUpdateModule,

    [switch]$NoModuleRequired,

    [switch]$PassThruSucceeded, 

    [bool]$NoUI = $false
    )

Process {

    $succeeded = $true

    try
    {

        Get-PSSession -Name $Service -ErrorAction SilentlyContinue  |
            ? {$_.State -ne [System.Management.Automation.Runspaces.RunspaceState]::Opened} |
            Remove-PSSession


        Get-PSSession -Name $Service -ErrorAction SilentlyContinue  |
            Select -Skip 1 |
            Remove-PSSession


        if (!($session = Get-PSSession -Name $Service -ErrorAction SilentlyContinue ))
        {

            switch -regex ($Service)
            {
                '(?i)^Premises(?<ContinentZone>EU|AA|AU|SA)?$'
                    {
                    $Credential = Get-BaoCredential -Service Premises
                    $Server = _GetCasExchangeServers -ContinentCode $Matches.ContinentZone | Get-Random
                    if (!$NoUI) {Write-Warning "Using Exchange server $Server"}
                    $Url = "http://$Server/powershell/"
                    }
                '(?i)^O365$'
                    {
                    $Credential = Get-BaoCredential -Service O365
#                    $Url = 'https://ps.outlook.com/powershell/'
                    $Url = 'https://outlook.office365.com/powershell-liveid/'
                    }
            }

            [Uri]$ConnectingURI = $null
            [Uri]::TryCreate($Url,[UriKind]::Absolute, [ref] $ConnectingURI) | Out-Null



            if ($Credential)
            {
                $secured = $ConnectingURI.Scheme -like '*s'
                $session = New-PSSession -Name $Service –ConfigurationName Microsoft.Exchange `
                    –ConnectionUri $ConnectingURI `
                    -Credential $Credential `
                    -Authentication $(if($secured){'basic'}else{'kerberos'}) `
                    –AllowRedirection `
                    -WarningAction SilentlyContinue `
                    -ErrorAction Stop
            }
            else { $succeeded = $false }
        }
        $session = $session | select -First 1

        if (!$NoModuleRequired)
        {

            if ( !(Get-Module $Service) -or $ForceUpdateModule )
            {

                if ( !(Get-Module $Service -ListAvailable) -or $ForceUpdateModule )
                {
                    Export-PSSession -Session $session `
                        -OutputModule "$PSHOME\Modules\$Service" `
                        -Force `
                        -AllowClobber `
                        -Encoding UTF8 `
                        -ErrorAction SilentlyContinue | Out-Null
                }

                Import-Module -Name $Service -Prefix $Service -Global -ErrorAction Stop -WarningAction SilentlyContinue

            }

            &(Get-Module $Service) Set-PSImplicitRemotingSession -PSSession $session # -createdByModule $True

        }

    }
    catch
    {
        $succeeded = $false
    }


    if ( $PassThruSucceeded ) { $succeeded }


}
}



function Connect-Accor
{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNull()]
        [ValidateSet('SharePointRoom', 'spREST', 'Premises', 'O365')]
        [string]$Service,

        [Parameter(Mandatory=$false, Position=1)]
        [ValidateSet('Room-0961', 'EU', 'AU', 'AA', 'SA', $null)]
        [string]$Detail,

        [switch]$PassThruSucceeded,

        [switch]$NoUI,

        [switch]$ForestScope
    )

    $succeeded = $true

    try {
        switch ( $Service )
        {
            'SharePointRoom'
            {
                if (!$Detail) {$Detail = 'Room-0961'}
                $succeeded = Connect-SharepointCSOM -Room $Detail  -PassThruSucceeded
            }
            'spREST'
            {
                if (!$Detail) {$Detail = 'Room-0961'}
                $succeeded = Connect-SharepointREST -Room $Detail -PassThruSucceeded
            }
            'O365'
            {
                $succeeded = Connect-Exchange -Service O365 -PassThruSucceeded -NoUI $NoUI   
            }
            'Premises'
            {
                $succeeded = Connect-Exchange -Service "Premises$Detail"  -PassThruSucceeded  -NoUI $NoUI  
                if ( $succeeded -and $ForestScope )
                {
                    $DCs = $AccorConfig.DomainsHash.Values | select -ExpandProperty DC
                    Set-PremisesADServerSettings -ViewEntireForest $true -SetPreferredDomainControllers $DCs -WarningAction SilentlyContinue
                }
            }
        }
    }
    catch { $succeeded = $false }
    
    if( $Host.UI.RawUI.WindowTitle -and !$NoUI )
    {

        $WhatDid = if ($succeeded) {'Connected To'} else {'Could not Connect To'}
        $WhatColor = if ($succeeded) {'Green'} else {'Red'}

        Write-Host "$WhatDid $Service service: $Detail" -ForegroundColor $WhatColor
    }

    if ( $PassThruSucceeded ) { $succeeded }

}
Export-ModuleMember -Function Connect-Accor

