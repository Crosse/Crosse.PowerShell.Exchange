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

function Add-RoomDelegate {
    param (
            [Parameter(Mandatory=$true)]
            [string]
            # The Identity of the room resource to which you want to add a delegate.
            $Identity, 

            [Parameter(Mandatory=$true)]
            [string]
            # The user who should be added as a room delegate.
            $Delegate
          );


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

    # Set the ResourceDelegates
    $resourceDelegates = (Get-CalendarProcessing -Identity $resource).ResourceDelegates
    if ( !($resourceDelegates.Contains((Get-User $Delegate).DistinguishedName)) ) {
        $resourceDelegates.Add( (Get-User $Delegate).DistinguishedName )
    }

    foreach ($i in 1..10) {
        $error.Clear()
        $resource | Set-CalendarProcessing -DomainController $DomainController `
                    -ResourceDelegates $resourceDelegates -ErrorAction SilentlyContinue
        if (![String]::IsNullOrEmpty($error[0])) {
            Write-Host -NoNewLine "."
            Start-Sleep $i
        } else {
            Write-Host "done."
            break
        }
    }    

    $resource | Set-Mailbox -DomainController $DomainController `
        -GrantSendOnBehalfTo $sobo

    <#
        .SYNOPSIS
        Adds a Delegate to a room resource in Microsoft Exchange.

        .DESCRIPTION
        Adds a Delegate to a room resource in Microsoft Exchange.  This entails
        granting Send-As and SendOnBehalfOf rights and adding the user to the
        resource's list of Resource Delegates.

        .INPUTS
        None.  Add-RoomDelegate does not accept any values from the pipeline.

        .OUTPUTS
        None.  Add-RoomDelegate does not return any values.
    #>
}
