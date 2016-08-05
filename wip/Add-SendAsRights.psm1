################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  Adds a user to a shared mailbox with Send-As and 
#               Send-On-Behalf-Of rights.
#
# 
# Copyright (c) 2009,2010 Seth Wright <wrightst@jmu.edu>
#
# Permission to use, copy, modify, and distribute this software for any
# purpose with or without fee is hereby granted, provided that the above
# copyright notice and this permission notice appear in all copies.
#
# THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
# WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
# MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
# ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
# WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
# ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
# OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
#
################################################################################

param ([string]$Identity, [string]$Delegate)

if ($Delegate -eq '' -or $Identity -eq '') {
    Write-Host "Please specify the Identity and Delegate"
    return
}

##################################
$DomainController = (Get-Content Env:\LOGONSERVER).Replace('\', '')
if ($DomainController -eq $null) { 
    Write-Warning "Could not determine the local computer's logon server!"
    return
}


$resource = Get-Mailbox $Identity
$delegate = Get-Mailbox $Delegate

if ($resource -eq $null) {
    Write-Error "Could not find Resource"
    return
}

Write-Host "Setting Permissions: "

# Grant Send-As rights to the delegate:
$Null = ($resource | Add-ADPermission -ExtendedRights "Send-As" -User $Delegate `
            -DomainController $DomainController)

# Grant SendOnBehalfOf rights to the delegate:
$sobo = (Get-Mailbox -DomainController $DomainController -Identity $resource).GrantSendOnBehalfTo
if ( !$sobo.Contains((Get-User $Delegate).DistinguishedName) ) {
    $null = $sobo.Add( (Get-User $Delegate).DistinguishedName )
    Write-Host "Current Users with Send rights:"
    $sobo |  Foreach-Object { $_.Name }
}
$resource | Set-Mailbox -DomainController $DomainController `
            -GrantSendOnBehalfTo $sobo

