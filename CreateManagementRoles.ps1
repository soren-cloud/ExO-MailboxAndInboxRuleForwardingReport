$Parent_RoleName = "Mail Recipients" # Name of the built-in role we are copying
$InboxRules_RoleName = "View-Only Inbox Rules" # Name of the new Inbox Rule role
$AcceptedDomains_RoleName = "View-Only Accepted Domains" # Name of the new Accepted Domains role

# Following commands will create a exact copy of the parent role
New-ManagementRole -Parent $Parent_RoleName -Name $InboxRules_RoleName 
New-ManagementRole -Parent $Parent_RoleName -Name $AcceptedDomains_RoleName

# Filter, so a list of all the unwanted entries (cmdlets) are stored in variables
# First - the Inbox Rules role
$InboxRules_EntriesToRemove = Get-ManagementRoleEntry "$InboxRules_RoleName\*" | `
	Where-Object {$_.Name -ne "Get-InboxRule" }

# Second - the Accepted Domains
$AcceptedDomains_Entries = Get-ManagementRoleEntry "$AcceptedDomains_RoleName\*" | `
	Where-Object {$_.Name -ne "Get-AcceptedDomain" }

# Now, remove the unwanted entries from the custom roles
# First - the Inbox Rules role
Foreach ($Entry in $AcceptedDomains_Entries)
	{
	$Role = $Entry.Identity
	$Name = $Entry.Name
	Remove-ManagementRoleEntry "$Role\$Name" -Confirm:$false 
	}

# Second - the Accepted Domains
Foreach ($Entry in $ManagementRoleEntries)
	{
	$Role = $Entry.Identity
	$Name = $Entry.Name
	Remove-ManagementRoleEntry "$Role\$Name" -Confirm:$false 
	}

# Verification
Get-ManagementRoleEntry "$InboxRules_RoleName\*"
Get-ManagementRoleEntry "$AcceptedDomains_RoleName\*"