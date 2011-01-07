################################################################################
# 
# $Id$
# 
# DESCRIPTION:  Creates a new resource in accordance with JMU's current naming
#               policies, etc.
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

param ( [string]$DisplayName, 
        [string]$Owner, 
        [string]$PrimarySmtpAddress = "")

$OU     = "ad.jmu.edu/ExchangeObjects/DistributionGroups"

##################################
$DomainController = (gc Env:\LOGONSERVER).Replace('\', '')
if ($DomainController -eq $null) { 
    Write-Warning "Could not determine the local computer's logon server!"
    return
}

if ( $DisplayName -eq '' -or $Owner -eq '') {
    Write-Output "-DisplayName and -Owner are required"
    return
}

$Owner = Get-Mailbox $Owner
if ($Owner -eq $null) {
    Write-Error "Could not find owner"
    return
}

$Name  = $DisplayName
$alias = $DisplayName
$alias = $alias.Replace(' ', '_')
$alias += "_DL"

$cmd  = "New-DistributionGroup -Name `"$Name`" -SamAccountName `"$alias`" -Type Security -Alias `"$alias`""
$cmd += "-DisplayName `"$DisplayName`" -ManagedBy `"$Owner`" -OrganizationalUnit `"$OU`" -DomainController $DomainController"

$error.Clear()

Invoke-Expression($cmd)

if (!([String]::IsNullOrEmpty($error[0]))) {
    return
}

$group = Get-DistributionGroup -DomainController $DomainController -Identity "$DisplayName"

if ( !$group) {
    Write-Output "Could not find $alias in Active Directory."
    return
}

#Start-Sleep 10
$group | Set-DistributionGroup -RequireSenderAuthenticationEnabled:$false -DomainController $DomainController
$group | Add-ADPermission -DomainController $DomainController `
            -AccessRights WriteProperty -Properties "Member" -User $Owner

