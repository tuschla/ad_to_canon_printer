param(
[string]$Gruppe)

Get-ADGroupMember $Gruppe -r | Get-ADUser -Properties DisplayName,EmailAddress,SAMAccountName,givenName, surName | sort givenName, surName 