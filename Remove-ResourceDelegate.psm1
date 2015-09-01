###########################################################################
# Copyright (c) 2009-2014 Seth Wright <wrightst@jmu.edu>
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
###########################################################################

function Remove-ResourceDelegate {
    [CmdletBinding(SupportsShouldProcess=$true,
            ConfirmImpact="High")]

    param (
            [Parameter(Mandatory=$true)]
            [Alias("Identity")]
            [ValidateNotNullOrEmpty()]
            [string]
            $ResourceIdentity,

            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Delegate,

            [switch]
            $RemoveFullMailboxAccess = $true,

            [switch]
            $RemoveSendAs = $true,

            [switch]
            $RemoveSendOnBehalfTo = $true,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            $DomainController
          )

    BEGIN {
        Write-Verbose "Performing initialization actions."

        if ([String]::IsNullOrEmpty($DomainController)) {
            $ForceRediscovery = [System.DirectoryServices.ActiveDirectory.LocatorOptions]::ForceRediscovery
            while ($dc -eq $null) {
                Write-Verbose "Finding a global catalog to use for this operation"
                $controller = [System.DirectoryServices.ActiveDirectory.Domain]::`
                    GetCurrentDomain().FindDomainController($ForceRediscovery)
                if ($controller.IsGlobalCatalog() -eq $true) {
                    Write-Verbose "Found $($controller.Name)"
                    $dc = $controller.Name
                } else {
                    Write-Verbose "Discovered domain controller $($controller.Name) is not a global catalog; reselecting"
                }
            }

            if ($dc -eq $null) {
                Write-Error "Could not find a domain controller to use for the operation."
                continue
            }
        } else {
            $dc = $DomainController
        }

        while ($resource -eq $null) {
            Write-Verbose "Resolving resource $ResourceIdentity into a mailbox"
            $resource = Get-Mailbox -Identity $ResourceIdentity `
                                    -DomainController $dc `
                                    -ErrorAction Stop
        }
    } # end 'BEGIN{}'

    # This section executes for each object in the pipeline.
    PROCESS {
        try {
            $objUser = Get-User $Delegate -DomainController $dc -ErrorAction Stop
        } catch {
            Write-Error -ErrorRecord $_
            return
        }

        $desc = 'Remove "{0}" from "{1}" as a delegate' -f $objUser.DisplayName, $resource.DisplayName
        $caption = $desc
        $warning = "Are you sure you want to perform this action?`n"
        $warning += 'This will remove {0} as a delegate of "{1}" ' -f $objUser.Name, $resource.DisplayName

        if (!$PSCmdlet.ShouldProcess($desc, $warning, $caption)) {
            Write-Error "User cancelled the operation."
            return
        }

        if ($RemoveSendAs) {
            Write-Verbose "Removing Send-As permission on $resource to $objUser"
            try {
                $null = Remove-ADPermission -Identity $resource.DistinguishedName `
                            -User $objUser.DistinguishedName `
                            -ExtendedRights "Send-As" `
                            -DomainController $dc `
                            -ErrorAction Stop `
                            -Confirm:$false
            } catch {
                Write-Error -ErrorRecord $_
            }
        }

        if ($RemoveFullMailboxAccess) {
            Write-Verbose "Removing FullAccess mailbox permission on $resource to $objUser"
            try {
                $null = Remove-MailboxPermission -Identity $resource `
                            -User $objUser.Identity `
                            -AccessRights FullAccess `
                            -DomainController $dc `
                            -ErrorAction Stop `
                            -Confirm:$false
            } catch {
                Write-Error -ErrorRecord $_
            }
        }

        # Loop through and remove any users who are no longer valid.
        # Do this always, instead of only when $RemoveSendOnBehalfTo is specified.
        $sobo = (Get-Mailbox -DomainController $dc -Identity $resource).GrantSendOnBehalfTo
        $dnsToRemove = @()
        if ($sobo.Count -gt 0) {
            Write-Verbose "Identifying invalid users in the GrantSendOnBehalfTo list"
            $dirty = $false
            foreach ($dn in $sobo) {
                $rtd = (Get-User -Identity $dn -ErrorAction SilentlyContinue).RecipientTypeDetails
                if ( $rtd -ne 'MailUser' -and $rtd -ne 'UserMailbox' ) {
                    $dnsToRemove += $dn
                    $dirty = $true
                }
            }
            foreach ($dn in $dnsToRemove) {
                Write-Verbose "Removing $dn from GrantSendOnBehalfTo list because user no longer has a mailbox"
                $null = $sobo.Remove($dn)
            }
        }

        if ($RemoveSendOnBehalfTo) {
            # Remove SendOnBehalfOf rights from the owner, if appropriate
            if ( $sobo.Contains($objUser.DistinguishedName) ) {
                $null = $sobo.Remove( $objUser.DistinguishedName )
                $dirty = $true
            }
        }

        if ($dirty) {
            Write-Verbose "Saving changes made to the GrantSendOnBehalfTo list"
            try {
                $null = Set-Mailbox -Identity $resource `
                        -GrantSendOnBehalfTo $sobo `
                        -DomainController $dc `
                        -ErrorAction Stop
            } catch {
                Write-Error -ErrorRecord $_
            }
        }

        if ($resource.RecipientTypeDetails -match "Room|Equipment") {
            Write-Verbose "Removing $objUser as a resource delegate on $resource"
            $resourceDelegates = (Get-CalendarProcessing -Identity $resource).ResourceDelegates
            if ( $resourceDelegates.Contains($objUser.DistinguishedName) ) {
                $null = $resourceDelegates.Remove($objUser.DistinguishedName)
            }

            try {
                $null = Set-CalendarProcessing -Identity $resource `
                            -ResourceDelegates $resourceDelegates `
                            -DomainController $dc `
                            -ErrorAction Stop
            } catch {
                Write-Error -ErrorRecord $_
            }
        }

        $resourceType = $resource.RecipientTypeDetails
    }
}
