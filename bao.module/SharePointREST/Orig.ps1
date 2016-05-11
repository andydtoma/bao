
function Invoke-SPORestMethod {
    [CmdletBinding()]
    [OutputType([int])]
    Param (
        # The REST endpoint URL to call.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.Uri]$Url,

        # Specifies the method used for the web request. The default value is "Get".
        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Get", "Head", "Post", "Put", "Delete", "Trace", "Options", "Merge", "Patch")]
        [string]$Method = "Get",

        # Additional metadata that should be provided as part of the Body of the request.
        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [object]$Metadata,

        # The "X-RequestDigest" header to set. This is most commonly used to provide the form digest variable. Use "(Invoke-SPORestMethod -Url "https://contoso.sharepoint.com/_api/contextinfo" -Method "Post").GetContextWebInformation.FormDigestValue" to get the Form Digest value.
        [Parameter(Mandatory = $false, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [string]$RequestDigest,
        
        # The "If-Match" header to set. Provide this to make sure you are not overwritting an item that has changed since you retrieved it.
        [Parameter(Mandatory = $false, Position = 4)]
        [ValidateNotNullOrEmpty()]
        [string]$ETag, 
        
        # To work around the fact that many firewalls and other network intermediaries block HTTP verbs other than GET and POST, specify PUT, DELETE, or MERGE requests for -XHTTPMethod with a POST value for -Method.
        [Parameter(Mandatory = $false, Position = 5)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Get", "Head", "Post", "Put", "Delete", "Trace", "Options", "Merge", "Patch")]
        [string]$XHTTPMethod,

        [Parameter(Mandatory = $false, Position = 6)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Verbose", "MinimalMetadata", "NoMetadata")]
        [string]$JSONVerbosity = "Verbose",

        # If the returned data is a binary data object such as a file from a SharePoint site specify the output file name to save the data to.
        [Parameter(Mandatory = $false, Position = 7)]
        [ValidateNotNullOrEmpty()]
        [string]$OutFile
    )

    Begin {
        if ((Get-Module Microsoft.Online.SharePoint.PowerShell -ListAvailable) -eq $null) {
            throw "The Microsoft SharePoint Online PowerShell cmdlets have not been installed."
        }
        if ($AccorConfig.SpoCred -eq $null) {
            [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null
            $cred = Get-BaoCredential -Service Room0961
            $AccorConfig.SpoCred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($cred.UserName, $cred.Password)
        } 

    }
    Process {
        $request = [System.Net.WebRequest]::Create($Url)
        $request.Credentials = $AccorConfig.SpoCred
        $odata = ";odata=$($JSONVerbosity.ToLower())"
        $request.Accept = "application/json$odata"
        $request.ContentType = "application/json$odata"   
        $request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")   
        $request.Method = $Method.ToUpper()

        if(![string]::IsNullOrEmpty($RequestDigest)) {
            $request.Headers.Add("X-RequestDigest", $RequestDigest)
        }
        if(![string]::IsNullOrEmpty($ETag)) {
            $request.Headers.Add("If-Match", $ETag)
        }
        if($XHTTPMethod -ne $null) {
            $request.Headers.Add("X-HTTP-Method", $XHTTPMethod.ToUpper())
        }
        if ($Metadata -is [string] -and ![string]::IsNullOrEmpty($Metadata)) {
            $body = [System.Text.Encoding]::UTF8.GetBytes($Metadata)
            $request.ContentLength = $body.Length
            try {
                $stream = $request.GetRequestStream()
                $stream.Write($body, 0, $body.Length)
            }
            finally { $stream.Dispose() }
        } elseif ($Metadata -is [byte[]] -and $Metadata.Count -gt 0) {
            $request.ContentLength = $Metadata.Length
            try {
                $stream = $request.GetRequestStream()
                $stream.Write($Metadata, 0, $Metadata.Length)
            }
            finally { $stream.Dispose() }
        } else {
            $request.ContentLength = 0
        }
 

        try {
            $response = $request.GetResponse()
        }
        catch [System.Management.Automation.MethodInvocationException] {
            if ($_.Exception.InnerException -is [System.Net.WebException]) {
                $exceptionResponse = $_.Exception.InnerException.Response
                $streamReader = New-Object System.IO.StreamReader $exceptionResponse.GetResponseStream()
                $data = $streamReader.ReadToEnd()
                $streamReader.Dispose()
                throw $data
            }
            else {    throw $_.Exception.InnerException }
        }
        catch {
            throw $_.Exception
        }

        try {
            $streamReader = New-Object System.IO.StreamReader $response.GetResponseStream()
            try {
                # If the response is a file (a binary stream) then save the file our output as-is.
                if ($response.ContentType.Contains("application/octet-stream")) {
                    if (![string]::IsNullOrEmpty($OutFile)) {
                        $fs = [System.IO.File]::Create($OutFile)
                        try {
                            $streamReader.BaseStream.CopyTo($fs)
                        } finally {
                            $fs.Dispose()
                        }
                        return
                    }
                    $memStream = New-Object System.IO.MemoryStream
                    try {
                        $streamReader.BaseStream.CopyTo($memStream)
                        Write-Output $memStream.ToArray()
                    } finally {
                        $memStream.Dispose()
                    }
                    return
                }
                # We don't have a file so assume JSON data.
                $data = $streamReader.ReadToEnd()

                # In many cases we might get two ID properties with different casing.
                # While this is legal in C# and JSON it is not with PowerShell so the
                # duplicate ID value must be renamed before we convert to a PSCustomObject.
                if ($data.Contains("`"ID`":") -and $data.Contains("`"Id`":")) {
                    $data = $data.Replace("`"ID`":", "`"ID-dup`":");
                }

                $results = ConvertFrom-Json -InputObject $data

                # The JSON verbosity setting changes the structure of the object returned.
                if ($JSONVerbosity -ne "verbose" -or $results.d -eq $null) {
                    Write-Output $results
                } else {
                    Write-Output $results.d 
                }
            } finally {
                $streamReader.Dispose()
            }
        } finally {
            $response.Dispose()
        }
    }
    End {
    }
}
Export-ModuleMember -Function Invoke-SPORestMethod

function Get-BaoSPBasic {
    [CmdletBinding()]
    [OutputType([int])]
    Param (
        # The REST endpoint URL to call.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.Uri]$Url,

        [Switch]$Unlimited,

        [Switch]$Metadata

    )
    Begin {
        if (!(Connect-Accor -Service spREST -NoUI -PassThruSucceeded))
        {
            throw 'Cannot connect to Sharepoint Room'
        }
    }
    Process {
        $u = $Url
        do
        {
            $r = Invoke-SPORestMethod -Url $u -Method Get
            $u = $r.__next

            if ($r.results -is [Array]) { $display = $r.results }
            else { $display = $r }

            $display |
            % {
                if ($md = $_.__metadata)
                {
                    $_.PSObject.Properties.Remove('__metadata') | Out-Null
                    if ($Metadata)
                    {
                        Add-Member -InputObject $_ -MemberType NoteProperty -Name __metadata -Value $md
                    }
                }
                $_
            }
        }
        while ($u -and $Unlimited)
    }
}
Export-ModuleMember -Function Get-BaoSPBasic

function Get-BaoSPList {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$List,

        [switch]$Metadata
    )
    Begin {
        $api = $AccorConfig.SharepointApiUrl
    }
    Process {
        if (!$AccorConfig.SpListsHash.ContainsKey($List)) {
            $AccorConfig.SpListsHash.$List = Get-BaoSPBasic -Url "$api/web/Lists/getbytitle('$List')" -Metadata
        }

        if ($Metadata) { $AccorConfig.SpListsHash.$List }
        else { $AccorConfig.SpListsHash.$List |  Select * -ExcludeProperty __metadata }

    }
}
Export-ModuleMember -Function Get-BaoSPList


function Get-BaoSPListFields {
    [CmdletBinding()]
    Param(
        [ValidateNotNullOrEmpty()]
        [string]$List,

        [switch]$Metadata

    )
    End {
        if (!$AccorConfig.SpListFieldsHash.ContainsKey($List) -and ($spList = Get-BaoSPList -List $List))
        {
            $oFilter = "`$Filter= CanBeDeleted"
            $oFilter+= " or InternalName eq 'Title'"
            $oFilter+= " or InternalName eq 'Created'"
            $oFilter+= " or InternalName eq 'Modified'"

            
            $oselect = '$Select= Title,InternalName,TypeAsString,Indexed,ReadOnlyField'
            $AccorConfig.SpListFieldsHash.$List = Get-BaoSPBasic -Url "$($spList.Fields.__deferred.uri)?$oFilter&$oSelect" -Unlimited -Metadata
        }
        if ($Metadata) { $AccorConfig.SpListFieldsHash.$List }
        else { $AccorConfig.SpListFieldsHash.$List |  Select * -ExcludeProperty __metadata }
    }

}
Export-ModuleMember -Function Get-BaoSPListFields

function Add-BaoSPListField {
   [CmdletBinding()]
    [OutputType([int])]
    Param (
        [ValidateNotNullOrEmpty()]
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$List,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 1)]
        [string]$FieldName,

        # Specifies the SP type of the field. The default value is "Text".
        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateSet('Text', 'Note', 'DateTime', 'Boolean', 'Number', 'Guid')]
        [string]$Type = 'Text'
    )
    Begin {
        $spList = Get-BaoSPList -List $List
    }
    Process{
        if (Get-BaoSPBasic -Url "$($spList.Fields.__deferred.uri)?`$Filter= InternalName eq '$FieldName'")
        {
            throw "$List has already a field named $FieldName"
        }
        else
        {
            switch ($Type)
            {
                'Note' {
                    $sType = 'MultiLineText'
                }
                'Boolean' {
                    $sType = [string]::Empty
                }
                Default {
                    $sType = $Type
                }
            }

            $spType = "SP.Field$sType"
            $spTypeKind = [Microsoft.SharePoint.Client.FieldType]::$Type.value__

            $body = @{
                __metadata = @{ type = $spType}
                Title = $FieldName
                FieldTypeKind = $spTypeKind
            } | ConvertTo-Json

            $f = Invoke-SPORestMethod -Url $spList.Fields.__deferred.uri -Method Post -Metadata $body -RequestDigest $AccorConfig.FormDigestValue 
        }
    }
    End {
        $oFilter = "`$Filter= CanBeDeleted or InternalName eq 'Title'"
        $oselect = '$Select= Title,InternalName,TypeAsString,Indexed,ReadOnlyField'
        $AccorConfig.SpListFieldsHash.$List = Get-BaoSPBasic -Url "$($spList.Fields.__deferred.uri)?$oFilter&$oSelect" -Unlimited -Metadata
    }
}
Export-ModuleMember -Function Add-BaoSPListField


function Add-BaoSPListFields {
   [CmdletBinding()]
    [OutputType([int])]
    Param (
        [ValidateNotNullOrEmpty()]
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$List,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 1)]
        [PSObject]$ObjectTemplate
    )
    Begin {
        $spList = Get-BaoSPList -List $List
    }

    Process{

        $ObjectTemplate |
        Confirm-SPLikeProperties |
        Get-Member -MemberType Properties |
        ? { $_.Name -notmatch 'PSComputerName|PSShowComputerName|RunspaceId|OriginatingServer|RemotePowerShellEnabled'} |
        %{
            $Prop = $_.Name
            if ($ObjectTemplate.$Prop -eq $null) { $sType = 'Text' }
            else {
                switch -regex ($ObjectTemplate.$Prop.GetType().FullName)
                {
                    '^System\.Collection|\[.*\]$' { $sType = 'Note'; break }
                    '^System\.String' { $sType = 'Text'; break }
                    '^System\.DateTime' { $sType = 'DateTime'; break }
                    '^System\.Boolean' { $sType = 'Boolean'; break }
                    '^System\.(U?Int|S?Byte|Float|Double|Decimal)' { $sType = 'Number'; break }
                    '^System\.Guid' { $sType = 'Guid'; break }
                    Default { $sType = 'Text' }
                }
            }

            try {
                Add-BaoSPListField -List $List -FieldName $_.Name -Type $sType
                Write-Output "Added $($_.Name) $sType"
            }
            catch {
            }
        }

    }
}
Export-ModuleMember -Function Add-BaoSPListFields


function Get-BaoSPListItems {
    [CmdletBinding()]
    Param(
        [ValidateNotNullOrEmpty()]
        [string]$List,

        [string]$Filter,

        [string[]]$Fields,

        [int]$PageSize=500,

        [switch]$Unlimited,

        [switch]$Metadata

    )
    End {
        if ($spList = Get-BaoSPList -List $List -Metadata)
        {
            $spFields = @(Get-BaoSPListFields -List $List | Select -ExpandProperty InternalName)
            if ($Fields)
            {
                $actualFields = @(Compare-Object $Fields $spFields -IncludeEqual -ExcludeDifferent |
                    Select -ExpandProperty InputObject)
            }
            else { $actualFields = $spFields }

            $pagedFilter = $false

            $oPageSize = "`$Top=$PageSize"

            $oFilter = if($Filter){"&`$Filter= $Filter"}else{[string]::Empty}

            $oSelect = if($Fields) {
                "&`$Select= $($actualFields -join ',')"
            } else {[string]::Empty}
            
            try {
                Get-BaoSPBasic -Url "$($spList.Items.__deferred.uri)?$oPageSize $oFilter $oSelect" -Unlimited:$Unlimited -Metadata:$Metadata
            }
            catch {
                if ($_.Exception.Message -match 'Microsoft\.SharePoint\.SPQueryThrottledException')
                {
                    $pagedFilter = $true
                }
                else
                {
                    throw $_.Exception.Message
                }
            }
            finally {
                if ($pagedFilter)
                {
                    $lastId = Get-BaoSPBasic -Url "$($spList.Items.__deferred.uri)?`$top=1&`$orderby=Id desc&`$select=Id" | Select -ExpandProperty Id
                    $currentId = 0
                    
                    while ($currentId -lt $lastId)
                    {
                        $nextId = $currentId + 5000

                        $oFilter = if($Filter){
                                    "&`$Filter=  ( $Filter ) and (Id gt $currentId) and (Id le $nextId)"
                                   }else{
                                    [string]::Empty
                                   }
                        Get-BaoSPBasic -Url "$($spList.Items.__deferred.uri)?$oPageSize $oFilter $oSelect" -Unlimited:$Unlimited -Metadata:$Metadata


                        $currentId = $nextId
                    }
                          
                }
             }

        }
    }

}
Export-ModuleMember -Function Get-BaoSPListItems

function Set-BaoSPBasic {
    [CmdletBinding()]
    [OutputType([int])]
    Param (
        # The REST endpoint URL to call.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        $Metadata,

        $UpdatedObject

    )
    Begin {
        if (!(Connect-Accor -Service spREST -NoUI -PassThruSucceeded))
        {
            throw 'Cannot connect to Sharepoint Room'
        }
    }
    Process {
        try {
            $body = [PSCustomObject]$UpdatedObject |
                Add-Member -MemberType NoteProperty -Name __metadata -Value @{type = $Metadata.type} -PassThru |
                ConvertTo-Json

            Invoke-SPORestMethod -Url $Metadata.uri -Method Post -XHTTPMethod Merge -ETag $Metadata.etag -Metadata $body -RequestDigest $AccorConfig.FormDigestValue
            
        }
        catch {
            throw $_.Exception.Message
        }
    }
}
Export-ModuleMember -Function Set-BaoSPBasic


$SPIdentities = @{
    baoMailbox = @{Guid='ExchangeGuid'; Upn='UserPrincipalName'; Alias = 'Alias'}
    zAccorPremisesUser = @{Guid='ObjectGuid'; Upn='UserPrincipalName'}
}

function Get-BaoSPListItem {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$List,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [string[]]$Fields,

        [switch]$Metadata


    )
    Begin {
        $spFields = Get-BaoSPListFields -List $List | Select -ExpandProperty InternalName

    }
    Process {
        switch ($Identity)
        {
            {$_ -as [Guid]} { $keyIdentity = $SPIdentities.$List.Guid; break }
            {$_ -match '^[\w\.-]+@[\w-]+(\.[\w-]+)+$'} { $keyIdentity = $SPIdentities.$List.Upn; break }
            {$_ -match '^[\w\.-]+$'} { $keyIdentity = $SPIdentities.$List.Alias; break }
            Default { $keyIdentity = $null }
        }

        if ($keyIdentity)
        {
            $item = Get-BaoSPListItems -List $List -Filter "$keyIdentity eq '$Identity'" -Fields $Fields -Metadata:$Metadata
            $item
        }
    }
}
Export-ModuleMember -Function Get-BaoSPListItem


function Set-BaoSPListItem {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$List,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        $Update
    )
    Begin {
        $spFields = Get-BaoSPListFields -List $List | Select -ExpandProperty InternalName

        $uObject = $Update

        $oFields = $uObject |
            Get-Member -MemberType Properties |
            Select -ExpandProperty Name |
            Compare-Object -ReferenceObject $spFields -IncludeEqual -ExcludeDifferent |
            select -ExpandProperty InputObject
    }
    Process {
        switch ($Identity)
        {
            {$_ -as [Guid]} { $keyIdentity = $SPIdentities.$List.Guid; break }
            {$_ -match '^[\w\.-]+@[\w-]+(\.[\w-]+)+$'} { $keyIdentity = $SPIdentities.$List.Upn; break }
            {$_ -match '^[\w\.-]+$'} { $keyIdentity = $SPIdentities.$List.Alias; break }
            Default { $keyIdentity = $null }
        }

        if ($keyIdentity)
        {
            $item = Get-BaoSPListItems -List $List -Filter "$keyIdentity eq '$Identity'" -Fields $keyIdentity -Metadata

            if ($Update -is [ScriptBlock])
            {
                $uObject = &$Update $Identity
                $oFields = $uObject |
                    Get-Member -MemberType Properties |
                    Select -ExpandProperty Name |
                    Compare-Object -ReferenceObject $spFields -IncludeEqual -ExcludeDifferent |
                    select -ExpandProperty InputObject
            }



            Set-BaoSPBasic `                -Metadata $item.__metadata `                -UpdatedObject (
                    $uObject |
                    Select -Property $oFields |
                    Convert-MultiValuedStringsToString
                 )
            
        }
    }
}
Export-ModuleMember -Function Set-BaoSPListItem


function Convert-MultiValuedStringsToString { 
    param 
    ( 
        [Parameter(Mandatory = $false,
                    ValueFromPipeline=$false, 
                    ValueFromPipelinebyPropertyName=$false, 
                    ValueFromRemainingArguments=$false, 
                    Position=0 
        )] 
        [ValidateNotNullOrEmpty()] 
        [String] $Seperator = "`n", 

        [Parameter(Mandatory = $true,  
                    ValueFromPipeline=$true, 
                    ValueFromPipelinebyPropertyName=$true, 
                    ValueFromRemainingArguments=$false, 
                    Position=1 
        )] 
        [ValidateNotNullOrEmpty()] 
        [object] $object 
    ) 
    Process { 
        $results= $object | 
        % { 
            $properties = New-Object PSObject     
            $_.PSObject.Properties |  
            % { 
                $propertyName = $_.Name 
                $propertyValue = $_.Value 
                If ($propertyValue -eq $null) {
                    Add-Member -inputObject $properties NoteProperty `                        -name $propertyName `                        -value $NULL 
                }
                elseif ($propertyValue.GetType().FullName -match '^System\.Collection|\[.*\]$' ) {  
                    $values = @() 
                    ForEach ($value In $propertyValue) { 
                        $values += $value.ToString() 
                    } 
                    Add-Member -inputObject $properties NoteProperty `                        -name $propertyName `                        -value "$([String]::Join($Seperator,$values))" 
                } Else {  
                    Add-Member -inputObject $properties NoteProperty `                        -name $propertyName `                        -value $propertyValue 
                } 
            } 
            $properties 
        } 
        return $results 
    } 
} 

filter Confirm-SPLikeProperties
{
    $SharepointSpecials = Write-Output Guid Id Title
    $InputObject = $_

    Select-Object -InputObject $InputObject -Property * -ExcludeProperty $SharepointSpecials |
    % {
        foreach ($special in $SharepointSpecials)
        {
            if ($InputObject.PSObject.Properties.Match($special).Count -gt 0)
            {
                Add-Member -InputObject $_ -MemberType NoteProperty -Name "Object$special" -Value $InputObject.$special
            }
        }
        $_
    }
}


function Add-BaoSPListItem {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$List,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        $Item
    )
    Begin {
        $spList = Get-BaoSPList -List $List -Metadata
        $itemsUri = $spList.Items.__deferred.uri

        $spFields = Get-BaoSPListFields -List $List | Select -ExpandProperty InternalName
    }
    Process {
        try {

            $oFields = $Item |
                Confirm-SPLikeProperties |
                Get-Member -MemberType Properties |
                Select -ExpandProperty Name |
                Compare-Object -ReferenceObject $spFields -IncludeEqual -ExcludeDifferent |
                select -ExpandProperty InputObject



            $body = $Item |
                Select -Property $oFields |
                Convert-MultiValuedStringsToString |
                Add-Member -MemberType NoteProperty -Name __metadata -Value @{type = $spList.ListItemEntityTypeFullName} -PassThru |
                ConvertTo-Json

            $response = Invoke-SPORestMethod -Url $itemsUri -Method Post -Metadata $body -RequestDigest $AccorConfig.FormDigestValue
            
        }
        catch {
            throw $_.Exception.Message
        }
    }
}
Export-ModuleMember -Function Add-BaoSPListItem


function Update-BaoSPListItem {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$List,

        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 1
        )]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        # This is an imperative delegate that finds THE source object for a given Identity
        # If null returned, than the List item is removed
        [ScriptBlock]$Source
    )
    Begin {
    }
    Process {
        $baseItem = Get-BaoSPListItem -List $List -Identity $Identity -Metadata -Fields Modified
        $sourceItem = &$Source $Identity

        if(!$baseItem) {
            if ($sourceItem) {
                Add-BaoSPListItem -List $List -Item ($sourceItem | Confirm-SPLikeProperties)
            }

        } else {
            if ($sourceItem) {
                if ( !($sM = $sourceItem.Modified -as [DateTime]) -or
                    ($sM -gt ($baseItem.Modified -as [DateTime]) )
                ) {
                    Set-BaoSPListItem -List $List -Identity $Identity -Update ($sourceItem | Confirm-SPLikeProperties)
                }
            } else {
                Remove-BaoSPListItem -List $List -Identity $Identity
            }
        }

    }
}
Export-ModuleMember -Function Update-BaoSPListItem

function Remove-BaoSPListItem {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$List,

        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 1
        )]
        [ValidateNotNullOrEmpty()]
        [string]$Identity

    )
    Begin {
    }
    Process {
        $baseItem = Get-BaoSPListItem -List $List -Identity $Identity -Metadata -Fields Modified

        if($baseItem) {
            try {
                Invoke-SPORestMethod -Url $baseItem.__metadata.uri -Method Post -XHTTPMethod Delete -ETag $baseItem.__metadata.etag -RequestDigest $AccorConfig.FormDigestValue
            }
            catch {
                throw $_.Exception.Message
            }
        }

    }
}
Export-ModuleMember -Function Remove-BaoSPListItem


