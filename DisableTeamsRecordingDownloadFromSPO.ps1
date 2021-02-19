param
(
    [string][Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$User,
    [switch][Parameter(Mandatory = $false)]$MFAUser,
    [string][Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()]$TeamDisplayName,
    [switch][Parameter(Mandatory = $false)]$CreateRecordingsFolder,
    [switch][Parameter(Mandatory = $false)]$ForceRemoveUserAsOwner
)

#region Checking if the required PowerShell modules are installed
$teamsModules = Get-Module -ListAvailable -Name MicrosoftTeams
if ($null -eq $teamsModules) {
    Write-Host "MicrosoftTeams module is not installed, installing it..."
    Install-Module MicrosoftTeams
}
$pnpModules = Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline
if ($null -eq $pnpModules) {
    Write-Host "SharePointPnPPowerShellOnline module is not installed, installing it..."
    Install-Module SharePointPnPPowerShellOnline -RequiredVersion 3.28.2012.0
}
#endregion Checking if the required PowerShell modules are installed

#region Functions
function GetAccessTokenPnP() {
    [cmdletbinding()]
    param
    (
        [PSCredential][Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()]$Credential,
        [switch][Parameter(Mandatory = $false)]$MFAUser,
        [string[]][Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$Scopes
    )
    if ($MFAUser) {
        $global:pnpGlobalConnection = Connect-PnPOnline -Scopes $Scopes
    }
    else {
        $global:pnpGlobalConnection = Connect-PnPOnline -Scopes $Scopes -Credentials $Credential    
    }
    return Get-PnPGraphAccessToken
}
function CheckCSVPermissionFolder() {
    param
    (
        [string][Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$FileName,
        [string][Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$ListName,
        [string][Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$Team
    )
    $CAMLQuery = "<View Scope='RecursiveAll'>
                    <Query>
                        <Where>
                            <Or>
                                <Eq>
                                    <FieldRef Name='LinkFilename'/>
                                    <Value Type='Text'>Recordings</Value>
                                </Eq>
                                <Eq>
                                    <FieldRef Name='LinkFilename'/>
                                    <Value Type='Text'>Registrazioni</Value>
                                </Eq>
                            </Or>
                        </Where>
                    </Query>
                </View>"
    $items = Get-PnPListItem -List $ListName -Query $CAMLQuery
    $ctx = Get-PnPContext
    $roleAssignments = $null
    foreach ($item in $items) {
        Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments | Out-Null
        $ctx.load($item.RoleAssignments)
        $ctx.load($item.Folder)
        $ctx.ExecuteQuery()
        $roleAssignments = $item.RoleAssignments
        foreach ($RoleAssignment in $roleAssignments) {
            $ctx.Load($RoleAssignment.Member)
            $ctx.Load($RoleAssignment.RoleDefinitionBindings)
            $ctx.ExecuteQuery()
            foreach ($RoleDefinition in ($RoleAssignment.RoleDefinitionBindings | Where-Object { $_.Hidden -eq $false })) {                  
                $RoleDefinition |
                Select-Object @{expression = { $Team }; Label = "Team" },
                @{expression = { $item.Folder.Name }; Label = "Folder" },
                @{expression = { $item.Folder.ServerRelativeUrl }; Label = "Relative path" },
                @{expression = { $RoleAssignment.Member.Title }; Label = "Role" },
                @{expression = { $_.Name }; Label = "Permission" } |
                Export-CSV $FileName -Append -Force -Encoding UTF8 -NoTypeInformation -Delimiter '|'
            }  
        }
    }	
}
function GetTeamWebsiteUrl() {
    param
    (
        [string][Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$TeamID,
        [string][Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$AccessToken
    )
    
    $graphUrl = "https://graph.microsoft.com/v1.0/groups/$TeamID/sites/root?`$select=webUrl"
    $headers = @{
        "Content-Type"  = "application/json"
        "Authorization" = "Bearer $AccessToken"
    }

    try {
        $response = Invoke-RestMethod -Uri $graphUrl -Headers $headers -Method Get -ContentType "application/json"
        return $response.webUrl
    }
    catch {
        return $null
    }
}
function ConvertFrom-SecondsSinceEpoch() {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)][string]$secondsSinceEpoch
    )

    $epoch = Get-Date -Year 1970 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    $utcDelta = New-TimeSpan -End ([datetime]::Now)  -Start ([datetime]::UtcNow)
    
    return $epoch.AddHours($utcDelta.Hours).AddSeconds($secondsSinceEpoch)
}
function ConvertFrom-Jwt {

    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Token,

        [Alias('ih')]
        [switch]$IncludeHeader
    )

    # Validate as per https://tools.ietf.org/html/rfc7519
    # Access and ID tokens are fine, Refresh tokens will not work
    if (!$Token.Contains(".") -or !$Token.StartsWith("eyJ")) { Write-Error "Invalid token" -ErrorAction Stop }

    # Extract header and payload
    $tokenheader, $tokenPayload = $Token.Split(".").Replace("-", "+").Replace("_", "/")[0..1]

    # Fix padding as needed, keep adding "=" until string length modulus 4 reaches 0
    while ($tokenheader.Length % 4) { Write-Debug "Invalid length for a Base-64 char array or string, adding ="; $tokenheader += "=" }
    while ($tokenPayload.Length % 4) { Write-Debug "Invalid length for a Base-64 char array or string, adding ="; $tokenPayload += "=" }

    Write-Debug "Base64 encoded (padded) header:`n$tokenheader"
    Write-Debug "Base64 encoded (padded) payoad:`n$tokenPayload"

    # Convert header from Base64 encoded string to PSObject all at once
    $header = [System.Text.Encoding]::ASCII.GetString([system.convert]::FromBase64String($tokenheader)) | ConvertFrom-Json
    Write-Debug "Decoded header:`n$header"

    # Convert payload to string array
    $tokenArray = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($tokenPayload))
    Write-Debug "Decoded array in JSON format:`n$tokenArray"

    # Convert from JSON to PSObject
    $tokobj = $tokenArray | ConvertFrom-Json
    Write-Debug "Decoded Payload:"

    if ($IncludeHeader) { $header }
    return $tokobj
}
#endregion Functions

#region Initializing variables
$roleName = "Read no download"
$roleDescription = "Can view pages and list items but can't download documents."
$spSiteUrl = $null
$spLibrary = $null
$members = $null
$visitors = $null
$membersRole = $null
$visitorsRole = $null
$logTime = Get-Date -Format "ddMMyyyy_HHmmss"
$start = Get-Date
$documentsListName = $null
$teams = $null
$accessToken = $null
$cred = $null
$teamOwners = $null
$shouldBeRemovedAsTeamOwner = $false
$pnpSiteConnection = $null
$teamChannels = $null
$channelFolder = $null
$channelRecFolder = $null
$accTokenDecoded = $null
$accTokenExpiration = $null
#endregion Initializing variables

#region Getting authenticated
try {
    if ($MFAUser) {
        $cred = Get-Credential -Message "Please enter your password:" -UserName $User
        $accessToken = GetAccessTokenPnP -Scopes @("Sites.Read.All") -MFAUser
        Connect-MicrosoftTeams
    }
    else {
        $cred = Get-Credential -Message "Please enter your password:" -UserName $User
        $accessToken = GetAccessTokenPnP -Scopes @("Sites.Read.All") -Credential $cred
        Connect-MicrosoftTeams -Credential $cred
    }
}
catch {
    Write-Error -Exception "Couldn't get authenticated. $($Exception.Message)"
    return
}
Disconnect-PnPOnline -Connection $global:pnpGlobalConnection
$accTokenDecoded = ConvertFrom-Jwt -Token $accessToken
$accTokenExpiration = ConvertFrom-SecondsSinceEpoch -secondsSinceEpoch $accTokenDecoded.exp
#endregion Getting authenticated

Start-Transcript -Path ".\Transcript_$logTime.txt"

# Getting the Team(s)
try {
    if ([string]::IsNullOrEmpty($TeamDisplayName)) {
        $teams = Get-Team
    }
    else {
        $teams = Get-Team -DisplayName $TeamDisplayName
    }
}
catch {
    Write-Error -Exception "Couldn't retrieve the Teams. $($Exception.Message)"
    return
}

foreach ($team in $teams) {
    Write-Host "***********************************************************"
    Write-Host "Processing Team: ""$($team.DisplayName)"""
    
    #region Handle user Team ownership
    $teamOwners = Get-TeamUser -Role Owner -GroupId $team.GroupId
    $shouldBeRemovedAsTeamOwner = $false
    if ($null -eq ($teamOwners | Where-Object { $_.User -eq $User })) {
        try {
            Write-Host "Adding user: ""$User"" as Team owner"
            Add-TeamUser -GroupId $team.GroupId -User $User -Role Owner
            $shouldBeRemovedAsTeamOwner = $true -and ($teamOwners.Count -ge 1)
        }
        catch {
            Write-Host "ERROR: Couldn't add user ""$User"" as owner of Team ""$($team.DisplayName)""." -ForegroundColor Red
            write-host "$_.Exception.Message" -ForegroundColor Red
            Write-Host "***********************************************************"
            continue
        }
    }
    else {
        Write-Host "User: ""$User"" is already a Team owner"
    }
    #endregion Handle user Team ownership
    
    $spSiteUrl = GetTeamWebsiteUrl -TeamID $team.GroupId -AccessToken $accessToken
    $pnpSiteConnection = $null
    if (![string]::IsNullOrEmpty($spSiteUrl)) {
        Write-Host "Team web site url: ""$spSiteUrl"""
        if ($MFAUser) {
            $pnpSiteConnection = Connect-PnPOnline -Url $spSiteUrl -UseWebLogin
        }
        else {
            $pnpSiteConnection = Connect-PnPOnline -Url $spSiteUrl -Credentials $cred
        }
        
        $web = Get-PnPWeb -Includes AssociatedMemberGroup, AssociatedVisitorGroup, AssociatedOwnerGroup
        
        #region Getting default Groups and their permissions 
        $members = $web.AssociatedMemberGroup
        $visitors = $web.AssociatedVisitorGroup
        $membersRole = Get-PnPGroupPermissions -Identity $members | Where-Object { $_.Hidden -eq $false }
        $visitorsRole = Get-PnPGroupPermissions -Identity $visitors | Where-Object { $_.Hidden -eq $false }
        Write-Host "Members group: ""$($members.Title)"" - Role: ""$($membersRole.Name)"""
        Write-Host "Visitors group: ""$($visitors.Title)"" - Role: ""$($visitorsRole.Name)"""
        #endregion Getting default Groups and their permission

        #region Handling custom permission level
        $readNoDownload = Get-PnPRoleDefinition | Where-Object { $_.Name -eq $roleName } | Select-Object -First 1
        if ($null -eq $readNoDownload) {
            Write-Host "Adding custom permission level ""$roleName"""
            $readNoDownload = Add-PnPRoleDefinition -RoleName $roleName -Description $roleDescription -Clone $visitorsRole.Name -Exclude OpenItems
        }
        else {
            $readNoDownload = $readNoDownload[0]
            Write-Host "Custom permission level ""$roleName"" already there"
        }
        #endregion Handling custom permission level
        
        #region Getting Documents document library
        $documentsListName = "Documents"
        $spLibrary = Get-PnPList -Identity $documentsListName
        if ($null -eq $spLibrary) {
            $documentsListName = "Documenti"
            $spLibrary = Get-PnPList -Identity $documentsListName
        }
        Write-Host "Document library: ""$($splibrary.Title)"""
        #endregion Getting Documents document library
        
        if ($CreateRecordingsFolder) {
            #region Creating Recordings folders in each Channel and setting the new permissions
            $teamChannels = Get-TeamChannel -GroupId $team.GroupId -MembershipType Standard
            foreach ($channel in $teamChannels) {
                $channelFolder = Get-PnPFolder -List $spLibrary | Where-Object { $_.Name -eq $channel.DisplayName } 
                if ($null -ne $channelFolder) {
                    $channelRecFolder = Get-PnPProperty -ClientObject $channelFolder -Property Folders 
                    $channelRecFolder = $channelRecFolder | Where-Object { $_.Name -eq "Recordings" -or $_.Name -eq "Registrazioni" }
                    if ($null -eq $channelRecFolder) {
                        write-host "Creating ""Recordings"" folder in channel folder ""$($channel.DisplayName)"""
                        $channelFolder.AddSubFolder("Recordings", $null)
                        $channelFolder.Context.ExecuteQuery()
                        $channelRecFolder = Get-PnPProperty -ClientObject $channelFolder -Property Folders 
                        $channelRecFolder = $channelRecFolder | Where-Object { $_.Name -eq "Recordings" -or $_.Name -eq "Registrazioni" }
                    }
                    else {
                        write-host """Recordings"" folder already present in channel folder ""$($channel.DisplayName)"""    
                    }
                    Write-Host "Changing permissions on folder ""$($channelRecFolder.ServerRelativeUrl)"""
                    Set-PnPFolderPermission -List $spLibrary -Identity $channelRecFolder.ServerRelativeUrl -Group $visitors -RemoveRole $visitorsRole.Name -AddRole $readNoDownload.Name
                    Set-PnPFolderPermission -List $spLibrary -Identity $channelRecFolder.ServerRelativeUrl -Group $members -RemoveRole $membersRole.Name -AddRole $readNoDownload.Name 
                }
                else {
                    Write-Host "WARNING: Unable to get channel folder for channel ""$($channel.DisplayName)""!" -ForegroundColor Yellow
                }
            }    
            #endregion Creating Recordings folders in each Channel and setting the new permissions
        }
        else {
            #region Changing permissions on Recordings folders
            $folders = Get-PnPFolder -List $spLibrary | Where-Object { $_.Name -eq "Recordings" -or $_.Name -eq "Registrazioni" } 
            Write-Host "Recordings folders found: $($folders.Count)"
            if ($folders.Count -gt 0) {
                $folders.ServerRelativeUrl
                foreach ($folder in $folders) {
                    Write-Host "Changing permissions on folder ""$($folder.ServerRelativeUrl)"""
                    Set-PnPFolderPermission -List $spLibrary -Identity $folder.ServerRelativeUrl -Group $visitors -RemoveRole $visitorsRole.Name -AddRole $readNoDownload.Name
                    Set-PnPFolderPermission -List $spLibrary -Identity $folder.ServerRelativeUrl -Group $members -RemoveRole $membersRole.Name -AddRole $readNoDownload.Name 
                }
            }
            #endregion Changing permissions on Recordings folders
        }

        CheckCSVPermissionFolder -FileName ".\ReportDisableTeamsRecordingDownloadFromSPO_$logTime.csv" -ListName $documentsListName -Team $team.DisplayName
        
        #region Handling User removal from Team Owners
        if ($shouldBeRemovedAsTeamOwner -or $ForceRemoveUserAsOwner) {
            Write-Host "Removing user: ""$User"" as Team owner"
            try {
                Remove-TeamUser -GroupId $team.GroupId -User $User
            }
            catch {
                Write-Error -Exception "Couldn't remove user '$User' from Team owners. $($Exception.Message)"
            }
        }
        else {
            Write-Host "Leaving user: ""$User"" as Team owner"
        }
        #endregion Handling User removal from Team Owners
    }
    else {
        Write-Host "ERROR: Unable to get team web site url!" -ForegroundColor Red
    }
    if ($null -ne $pnpSiteConnection) {
        Disconnect-PnPOnline -Connection $pnpSiteConnection
    }
    Write-Host "***********************************************************"

    #region EVALUATING IF THE ACCESS TOKEN IS ABOUT TO EXPIRE
    if ((Get-Date) -gt $accTokenExpiration.AddMinutes(-5)) {
        Write-Host "Access token is expired, requesting a new access token" -ForegroundColor Magenta
        if ($MFAUser) {
            $accessToken = GetAccessTokenPnP -Scopes @("Sites.Read.All") -MFAUser  
        }
        else {
            $accessToken = GetAccessTokenPnP -Scopes @("Sites.Read.All") -Credential $cred
        }
        $accTokenDecoded = ConvertFrom-Jwt -Token $accessToken
        $accTokenExpiration = ConvertFrom-SecondsSinceEpoch -secondsSinceEpoch $accTokenDecoded.exp
    }
    #endregion
}
$end = Get-Date
Write-Host -ForegroundColor Green ("Script completed in {0:g}" -f ($end - $start))
Write-Host "Recordings folders permissions report saved at ""$(Get-Location)\ReportDisableTeamsRecordingDownloadFromSPO_$logTime.csv"""
Disconnect-MicrosoftTeams
Stop-Transcript