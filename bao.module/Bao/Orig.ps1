
$AccorConfig = [PSCustomObject]@{
    Store = "$PSScriptRoot\Config"
    Logs = "$PSScriptRoot\Logs"
    CredentialHash = @{}
    DomainsHash = $null
    SharepointApiUrl = $null
    SpoCred = $null
    FormDigestValue = $null
    FormDigestExpireAt = $null
    SpListsHash = @{}
    SpListFieldsHash = @{}
    SharePointClientContext = $null
    SharePointListPageSize = 150
    StringMarks = ((0..99 | % { "h{0:D2}" -f $_}) + (Write-Output a c e g i k m n p r t v x z)) | sort
    ExecutionServer = 's-eu-evy01mxs11.eu.accor.net'
}

function Get-BaoConfiguration
{
    $AccorConfig
}
Export-ModuleMember -Function Get-BaoConfiguration

function Register-BaoCredential
{
    Param (
        [Parameter(Position=0,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Room0961', 'Premises', 'O365', 'ScheduledJob')]
        [string]$Service
    )
    Begin
    {

        $credBaoxml = "$($AccorConfig.Store)\Credentials.baoxml"
        $credHash = $null

        if ( Test-Path -Path $credBaoxml -PathType Leaf )
        {
            $credHash = Import-Clixml -Path $credBaoxml -ErrorAction SilentlyContinue
        }
        if ( $credHash -isnot [HashTable] ) { $credHash = @{} }

    }
    Process {

        if ($oldCred = (Get-BaoCredential -Service $Service)  -as [PSCredential])
        {
            $oldUser = $oldCred.UserName
        }

        $newCred = Get-Credential -UserName $oldUser -Message "Please, Enter Credentials for $Service"

        $credService = "$env:USERDNSDOMAIN;$env:USERNAME;$Service"

        $AccorConfig.CredentialHash.$credService = $newCred

        $byteKey =  [Text.Encoding]::UTF8.GetBytes($credService.PadRight(24, '~').ToUpper()) | Select -First 24
        $newBaoCred = [PSCustomObject]@{
            UserName = $newCred.UserName
            Password = ConvertFrom-SecureString -SecureString $newCred.Password -Key $byteKey
        }
        
        $credHash.$credService = $newBaoCred
        
    }
    End
    {
        $storeDir = New-Item -ItemType Directory -Path $AccorConfig.Store -Force
        
        $credHash | Export-Clixml -Path $credBaoxml -Encoding UTF8 -Depth 4

    }
}
Export-ModuleMember -Function Register-BaoCredential

function Get-BaoCredential
{
    Param (
        [Parameter(Position=0)]
        [AllowNull()]
        [ValidateSet('Room0961', 'Premises', 'O365', 'ScheduledJob')]
        [string]$Service
    )

    $credService = "$env:USERDNSDOMAIN;$env:USERNAME;$Service"

    if ($Service -and  ($cred = $AccorConfig.CredentialHash.$credService))
    {
        $cred
    }
    else
    {
        $credBaoxml = "$($AccorConfig.Store)\Credentials.baoxml"
        if ( Test-Path -Path $credBaoxml -PathType Leaf )
        {
            $credHash = Import-Clixml -Path $credBaoxml -ErrorAction SilentlyContinue
            if ( $credHash -is [HashTable] )
            {
                if ($Service)
                {
                    if (($baoCred = $credHash.$credService) -and $baoCred.UserName -and $baoCred.Password)
                    {
                        $byteKey =  [Text.Encoding]::UTF8.GetBytes($credService.PadRight(24, '~').ToUpper()) | Select -First 24
                        
                        #Out-BaoLog -Message "CredMessage = $credService Key = $byteKey"
                        
                        $ss = ConvertTo-SecureString -String $baoCred.Password -Key $byteKey

                        $cred = New-Object PSCredential ($baoCred.UserName, $ss)

                        $AccorConfig.CredentialHash.$credService = $cred
                        $cred
                    }
                }
                else
                {
                    $credHash.Keys |
                    ? { $_ -match "^$env:USERDNSDOMAIN;$env:USERNAME;" } |
                    %{
                        [PSCustomObject]@{
                            Service = $_ -split ';' | Select -Last 1
                            CredentialUserName = $credHash.$_.UserName
                        }
                    }
                }
            }
        }
    }
}
Export-ModuleMember -Function Get-BaoCredential

function Out-BaoLog
{
    Param
    (
        [Parameter(ValueFromPipeline=$true, Position = 0)]
        [string]$Message
    )
    Begin
    {
        if ( !(Test-Path $AccorConfig.Logs -PathType Container) )
        {
            New-Item -ItemType Directory -Path $AccorConfig.Logs -Force | Out-Null
        }

        $CallingFunction = Get-PSCallStack |
            Select -Skip 1 |
            ? { $_.Command } |
            Select -First 1 |
            % { $_.Command -replace '[<>]', [string]::Empty }
            #Select -ExpandProperty Command

        $LogFile = 'LogBao-{1}-{0:yyyy-MM-dd}.log' -f (Get-Date).Date , $CallingFunction
        $LogFilePath = "$($AccorConfig.Logs)\$LogFile"



    }
    Process
    {

        $line = "{0:yyyy.MM.dd-HH:mm:ss.fff} {1,-30} {2}" -f
                    (Get-Date),
                    $CallingFunction,
                    $Message


        $retries = 1000
#        do
#        {
#            try
#            {
#                "{0:yyyy.MM.dd-HH:mm:ss.fff} {1,-30} {2}" -f
#                    (Get-Date),
#                    $CallingFunction,
#                    $Message |
#                Add-Content -Path "$($AccorConfig.Logs)\$LogFile" -Force
#
#                $written = $true
#            }
#            catch
#            {
#                $written = $false
#                Start-Sleep -Milliseconds 30
#            }
#        }
#        until ($written -or ($retries-- -le 0))
        while ($retries-- -gt 0)
        {
            try
            {
                [IO.File]::OpenWrite($LogFilePath).Close()
                $line | Add-Content -Path $LogFilePath
                break
            }
            catch { }
        }
    }
    
}
Export-ModuleMember -Function Out-BaoLog

function Out-BaoWorkflowLog
{
    Param
    (
        [string]$Caller,
        [string]$Context,
        [string]$Message
    )
    Begin
    {
        if ( !(Test-Path $AccorConfig.Logs -PathType Container) )
        {
            New-Item -ItemType Directory -Path $AccorConfig.Logs -Force | Out-Null
        }


        $LogFile = 'LogWorkflow-{0:yyyy-MM-dd}.csv' -f (Get-Date).Date 
        $LogFilePath = "$($AccorConfig.Logs)\$LogFile"



    }
    Process
    {

        $line = [PSCustomObject][ordered]@{
                    TimeStamp = "{0:yyyy.MM.dd-HH:mm:ss.fff}" -f  (Get-Date)
                    Caller = $Caller
                    Context = $Context
                    Message = $Message
                }

        try {
            $retry = $false
            $line | Export-Csv $LogFilePath -Append -Delimiter ',' -Encoding UTF8 -NoTypeInformation
        }
        catch {
            $retry = $true
        }
        finally {
            if ($retry) {
                Start-Sleep -Milliseconds 100
                $line | Export-Csv $LogFilePath -Append -Delimiter ',' -Encoding UTF8 -NoTypeInformation
            }
        }
    }
    
}
Export-ModuleMember -Function Out-BaoWorkflowLog


function Register-DomainsHash
{
    $domainsHash = @{}

    Get-ADForest |
    Select -ExpandProperty Domains |
    %{
        $dom = $_
        $dc = (Get-ADDomainController -Discover -DomainName $dom -NextClosestSite).HostName | Get-Random
        $nebiosName = (Get-ADDomain -Identity $dom).NetBIOSName
        $domainsHash[$dom] = [PSCustomObject]@{
                DC = $dc
                NetBIOSName = $nebiosName
            }
    }

    $domBaoxml = "$($AccorConfig.Store)\Domains.baoxml"
    $domainsHash | Export-Clixml -Path $domBaoxml -Encoding UTF8 -Depth 4
    $AccorConfig.DomainsHash = $domainsHash
    $domainsHash
}

function Get-DomainsHash
{
    if (!($domainsHash = $AccorConfig.DomainsHash))
    {
        $domBaoxml = "$($AccorConfig.Store)\Domains.baoxml"
        if ( Test-Path -Path $domBaoxml -PathType Leaf )
        {
            $domainsHash = Import-Clixml -Path $domBaoxml -ErrorAction SilentlyContinue
            if ( $domainsHash -isnot [HashTable] ) { $domainsHash = $null }
        }

        if (!$domainsHash) { $domainsHash = Register-DomainsHash }

        $AccorConfig.DomainsHash = $domainsHash        

    }
    $domainsHash
}


function Initialize-Module
{
    $Error.Clear()
    $StorePath = [string]$AccorConfig.Store

    if (Test-Path -Path $StorePath -PathType Container)
    {
        $AccorConfig.Store = Get-Item -Path $StorePath -ErrorAction Stop
    }
    else
    {
        try
        {
            $AccorConfig.Store = New-Item -ItemType Directory -Path $StorePath -ErrorAction Stop
        }
        catch [System.IO.IOException]
        {
            $AccorConfig.LastError = $_.CategoryInfo
            exit
           
        }
        finally
        {
            $Error.Clear()
        }
    }

    Get-DomainsHash | Out-Null

}

Initialize-Module


