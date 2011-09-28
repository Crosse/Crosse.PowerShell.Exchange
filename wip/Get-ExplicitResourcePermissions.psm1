################################################################################
# 
# $Id$
# 
# DESCRIPTION:  Gets mailbox permissions, removing extraneous entries.
#               Will NOT return inherited ACEs.
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

param ( $Identity=$null, 
        [switch]$MailboxPermissions=$false, 
        [switch]$SendAsPermissions=$false,
        [switch]$SendOnBehalfPermissions=$false,
        $inputObject=$null)

# This section executes only once, before the pipeline.
BEGIN {
    if ($inputObject) {
        Write-Output $inputObject | &($MyInvocation.InvocationName)
        break
    }

} # end 'BEGIN{}'

# This section executes for each object in the pipeline.
PROCESS {
    if ($_) { $Identity = $_ }

    if ($Identity -eq $null) {
        Write-Error "No mailbox specified"
        return
    } else {
        $Identity = Get-Mailbox $Identity -ErrorAction SilentlyContinue

        if ($Identity -eq $null) {
            Write-Error "$Identity is not a mailbox."
            return
        }
    }

    if ($MailboxPermissions -eq $false -and `
            $SendAsPermissions -eq $false -and `
            $SendOnBehalfPermissions -eq $false) {

        $MailboxPermissions = $true
        $SendAsPermissions = $true
        $SendOnBehalfPermissions = $true
    }

    if ($MailboxPermissions -eq $true) {
        $perms = Get-MailboxPermission $Identity.ToString() | ? { `
                $_.IsInherited -eq $False -and $_.AccessRights -eq "FullAccess" } 

        $perms | 
            Select @{Name="Identity"; Expression={ $_.Identity.Name } }, `
            User, `
            @{Name="Right"; Expression={ $_.AccessRights } }
    }

    if ($SendAsPermissions -eq $true) {
        $perms = Get-ADPermission $Identity.Identity | ? { `
            $_.Deny -eq $false -and `
            $_.User -notmatch "SELF" -and `
            $_.ExtendedRights -match "Send-As" }

        $perms | 
            Select @{Name="Identity"; Expression={ $_.Identity.Name } }, `
                User, `
                @{Name="Right"; Expression={ $_.ExtendedRights } }
    }

    if ($SendOnBehalfPermissions -eq $true) {
        $Identity.GrantSendOnBehalfTo |
            Select @{Name="Identity"; Expression={ $Identity.Name } }, `
            @{Name="User"; Expression={ $_.Name } }, `
            @{Name="Right"; Expression={ "SendOnBehalf" } }
    }
}
