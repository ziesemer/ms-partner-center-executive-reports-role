# Mark A. Ziesemer, www.ziesemer.com - 2025-01-19, 2025-12-21

#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess)]
Param(
	[Parameter(Mandatory)]
	[guid]$tenantId,
	[Parameter(Mandatory)]
	[guid]$appId,
	[Parameter(Mandatory)]
	[string[]]$userId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

function Write-Log{
	# False-positive
	[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '')]

	[CmdletBinding()]
	Param(
		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[object]$Message,

		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('ERROR', 'WARN', 'INFO', 'DEBUG', 'TRACE', IgnoreCase=$false)]
		[string]$Severity = 'INFO'
	)

	if($Severity -ceq 'TRACE'){
		$color = [ConsoleColor]::DarkGray
	}elseif($Severity -ceq 'DEBUG'){
		$color = [ConsoleColor]::Gray
	}elseif($Severity -ceq 'INFO'){
		$color = [ConsoleColor]::Cyan
	}elseif($Severity -ceq 'WARN'){
		$color = [ConsoleColor]::Yellow
	}elseif($Severity -ceq 'ERROR'){
		$color = [ConsoleColor]::Red
	}

	$msg = "$(Get-Date -f s) [$Severity] $Message"

	Write-Information ([System.Management.Automation.HostInformationMessage]@{
		Message = $msg
		ForegroundColor = $color
	})
}

function Invoke-ERAAuthorize{
	[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
	Param()
	$req = Invoke-RestMethod -Method Post `
		-Uri ('https://login.microsoftonline.com/' `
			+ $tenantId `
			+ '/oauth2/v2.0/devicecode') `
		-Body @{
			client_id = $appId
			scope = 'https://api.partnercenter.microsoft.com/user_impersonation offline_access'
		}

	Write-Host $req.message
	[void](Read-Host 'Press Enter once authentication is complete.')
	return $req
}

function Invoke-ERAAuthToken($code){
	$resp = Invoke-RestMethod -Uri ('https://login.microsoftonline.com/' `
			+ $tenantId `
			+ '/oauth2/v2.0/token') `
		-Method Post `
		-Body @{
			'client_id' = $appId
			'grant_type' = 'urn:ietf:params:oauth:grant-type:device_code'
			'device_code' = $code.device_code
		}
	$resp
}

function Invoke-ERAAuthTokenFromRefresh($token, $scope){
	$resp = Invoke-RestMethod -Uri ('https://login.microsoftonline.com/' `
			+ $tenantId `
			+ '/oauth2/v2.0/token') `
		-Method Post `
		-Body @{
			'client_id' = $appId
			'grant_type' = 'refresh_token'
			'refresh_token' = $token.refresh_token
			'scope' = $scope
		}
	$resp
}

# The Partner Center APIs are returning a UTF-8 BOM, along with a response header of:
#   Content-Type: application/json; charset=utf-8
# Even though the actual HTTP response contains the proper UTF-8 BOM (0xEF,0xBB,0xBF),
#   this is somehow being received through Invoke-RestMethod / Invoke-WebRequest as 0xFEFF.
# - https://github.com/PowerShell/PowerShell/issues/5007
# - https://windowsserver.uservoice.com/forums/301869-powershell/suggestions/11088009-invoke-webrequest-ignores-content-encoding
# - https://stackoverflow.com/questions/20388562/invoke-restmethod-not-recognizing-xml-with-byte-order-mark-served-by-sharepoint
# - https://stackoverflow.com/questions/47908382/convertfrom-json-invalid-json-primitive-%C3%AF
function Convert-PCFixEncoding($x){
	if($x[0] -eq [char]0xFEFF){
		$x = $x.Substring(1)
	}
	$x | ConvertFrom-Json
}

function Get-ERAUser($token, $userId){
	# Otherwise considered https://api.partnercenter.microsoft.com/v1/users/
	# - Documentation is effectively non-existent, only works by userId (GUID), not UserPrincipalName (UPN).

	# - https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
	$resp = Invoke-WebRequest -Uri ('https://graph.microsoft.com/v1.0/users/' `
			+ $userId) `
		-Headers @{
			'Authorization' = 'Bearer ' + $token.access_token
			'Accept' = 'application/json'
		}
	Convert-PCFixEncoding $resp.Content
}

function Get-ERARoles{
	[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', 'Get-ERARoles')]
	Param(
		$token
	)
	# - https://learn.microsoft.com/en-us/rest/api/partner-center/manage-account-profiles/get-all-available-roles
	$resp = Invoke-WebRequest -Uri 'https://api.partnercenter.microsoft.com/v1/roles' `
		-Headers @{
			'Authorization' = 'Bearer ' + $token.access_token
			'Accept' = 'application/json'
		}
	Convert-PCFixEncoding $resp.Content
}

function Get-ERARoleMember{
	Param(
		$token, $role
	)
	# TODO: Implement paging - though available documentation on this is not exactly clear how to do so.
	# - https://learn.microsoft.com/en-us/rest/api/partner-center/manage-account-profiles/get-user-members-by-role
	$resp = Invoke-WebRequest -Uri ('https://api.partnercenter.microsoft.com/v1/roles/' `
			+ $role.id `
			+ '/usermembers') `
		-Headers @{
			'Authorization' = 'Bearer ' + $token.access_token
			'Content-Type' = 'application/json'
			'Accept' = 'application/json'
		} `
		-Method Get
	Convert-PCFixEncoding $resp.Content
}

function Remove-ERARoleMember{
	[CmdletBinding(SupportsShouldProcess)]
	Param(
		$token, $role, $user
	)
	if(!$PSCmdlet.ShouldProcess($user.userPrincipalName)){
		return
	}
	# - https://learn.microsoft.com/en-us/rest/api/partner-center/manage-account-profiles/delete-user-member-from-role
	# As of 2025-12-21, this is failing with:
	# 	{
	# 	  "code": 2000,
	# 	  "description": "Account Id has to be set.",
	# 	  "data": [],
	# 	  "source": "PartnerFD"
	# 	}
	# ... even though there is no mention of "account" is listed in the above-documented API.
	# Looks like the Partner Center web UI instead issues a PATCH request through the otherwise internal and undocumented:
	# - https://partner.microsoft.com/en-us/dashboard/account/v3/api/authv2/user/...
	$resp = Invoke-WebRequest -Uri ('https://api.partnercenter.microsoft.com/v1/roles/' `
			+ $role.id `
			+ '/usermembers/' `
			+ $user.id) `
		-Headers @{
			'Authorization' = 'Bearer ' + $token.access_token
			'Content-Type' = 'application/json'
			'Accept' = 'application/json'
		} `
		-Method Delete
	Convert-PCFixEncoding $resp.Content
}

function Set-ERARoleMember{
	[CmdletBinding(SupportsShouldProcess)]
	Param(
		$token, $role, $user
	)
	if(!$PSCmdlet.ShouldProcess($user.userPrincipalName)){
		return
	}
	# - https://learn.microsoft.com/en-us/rest/api/partner-center/manage-account-profiles/add-new-user-member-to-role
	$resp = Invoke-WebRequest -Uri ('https://api.partnercenter.microsoft.com/v1/roles/' `
			+ $role.id `
			+ '/usermembers') `
		-Headers @{
			'Authorization' = 'Bearer ' + $token.access_token
			'Content-Type' = 'application/json'
			'Accept' = 'application/json'
		} `
		-Method Post `
		-Body (@{
			'accountId' = $tenantId
			'displayName' = $user.displayName
			'id' = $user.id
			'roleId' = $role.id
			'userPrincipalName' = $user.userPrincipalName
		} | ConvertTo-Json)
	Convert-PCFixEncoding $resp.Content
}

# This also fixes false-positive PSReviewUnusedParameter reports on these variables.
Write-Log "Connecting: tenantId=$tenantId, appId=$appId ..." -Severity 'DEBUG'

# - https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code

#   1) Calls https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode
Write-Log '/oauth2/v2.0/devicecode ...'
$authResponse = Invoke-ERAAuthorize

#   2) Calls https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
Write-Log '/oauth2/v2.0/token ...'
$token = Invoke-ERAAuthToken $authResponse

Write-Log 'Getting userReadToken ...'
$userReadToken = Invoke-ERAAuthTokenFromRefresh $token 'https://graph.microsoft.com/User.Read'

Write-Log 'Getting roles...'
$pcRoles = Get-ERARoles $token

$ervRole = $pcRoles.items | Where-Object{$_.name -eq 'Executive Report Viewer'}
Write-Log ('Found role: ' + ($ervRole | ConvertTo-Json -Compress))

$ervMembers = Get-ERARoleMember $token $ervRole
Write-Log ('Found role members: ' + ($ervMembers | ConvertTo-Json -Compress -Depth 4))

$rvRole = $pcRoles.items | Where-Object{$_.name -eq 'Report Viewer'}
Write-Log ('Found role: ' + ($rvRole | ConvertTo-Json -Compress))

$rvMembers = Get-ERARoleMember $token $rvRole
Write-Log ('Found role members: ' + ($rvMembers | ConvertTo-Json -Compress -Depth 4))

foreach($uid in $userId){
	Write-Log ('Getting user details: ' + $uid)
	$user = Get-ERAUser $userReadToken $uid

	if($uid -in $rvMembers.items.id){
		Write-Log '  - User already also in "Report Viewer" role, removing...'
		# $setRoleResp = Remove-ERARoleMember $token -role $ervRole -user $user
		# Write-Log ('  - Role response: ' + ($setRoleResp | ConvertTo-Json -Compress))
		Write-Log -Severity WARN '  Delete API is not working per documented specification, user must be removed from role manually.'
		Write-Log -Severity WARN '  Remove all Report Viewer-related roles from the Partner Center web UI, and reattempt.'
	}

	Write-Log '  - Adding "Executive Report Viewer" role ...'
	$setRoleResp = Set-ERARoleMember $token -role $ervRole -user $user
	Write-Log ('  - Role response: ' + ($setRoleResp | ConvertTo-Json -Compress))
}

Write-Log 'Done!'
