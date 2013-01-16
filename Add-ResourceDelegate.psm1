###########################################################################
# Copyright (c) 2009-2013 Seth Wright <wrightst@jmu.edu>
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

function Add-ResourceDelegate {
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
            $GrantFullMailboxAccess = $true,

            [switch]
            $GrantSendAs = $true,

            [switch]
            $GrantSendOnBehalfTo = $true,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            $SmtpServer = "mailgw.jmu.edu",

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            $From = "Exchange System <it-exmaint@jmu.edu>",

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string[]]
            $Bcc = @("wrightst@jmu.edu", "najdziav@jmu.edu", "eckardsl@jmu.edu"),

            [switch]
            $EmailDelegate=$true,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [ValidateScript({ (Test-Path $_) })]
            [string]
            $SharedMailboxTemplateEmail = (Join-Path $PSScriptRoot "SharedMailboxDelegateTemplateEmail.html"),

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [ValidateScript({ (Test-Path $_) })]
            [string]
            $ResourceMailboxTemplateEmail = (Join-Path $PSScriptRoot "ResourceMailboxDelegateTemplateEmail.html"),

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

        $desc = 'Add "{0}" to "{1}" as a delegate' -f $objUser.DisplayName, $resource.DisplayName
        $caption = $desc
        $warning = "Are you sure you want to perform this action?`n"
        $warning += 'This will make {0} a delegate of "{1}" ' -f $objUser.Name, $resource.DisplayName

        if (!$PSCmdlet.ShouldProcess($desc, $warning, $caption)) {
            Write-Error "User cancelled the operation."
            return
        }

        if ($GrantSendAs) {
            Write-Verbose "Granting Send-As permission on $resource to $objUser"
            try {
                $null = Add-ADPermission -Identity $resource.DistinguishedName `
                            -User $objUser.DistinguishedName `
                            -ExtendedRights "Send-As" `
                            -DomainController $dc `
                            -ErrorAction Stop
            } catch {
                Write-Error -ErrorRecord $_
            }
        }

        if ($GrantFullMailboxAccess) {
            Write-Verbose "Granting FullAccess mailbox permission on $resource to $objUser"
            try {
                $null = Add-MailboxPermission -Identity $resource `
                            -User $objUser.Identity `
                            -AccessRights FullAccess `
                            -AutoMapping:$false `
                            -DomainController $dc `
                            -ErrorAction Stop
            } catch {
                Write-Error -ErrorRecord $_
            }
        }

        # Loop through and remove any users who are no longer valid.
        # Do this always, instead of only when $GrantSendOnBehalfTo is specified.
        $sobo = (Get-Mailbox -DomainController $dc -Identity $resource).GrantSendOnBehalfTo
        Write-Verbose "Identifying invalid users in the GrantSendOnBehalfTo list"
        $dirty = $false
        foreach ($dn in $sobo) {
            if ( (Get-Mailbox -Identity $dn -ErrorAction SilentlyContinue) -eq $null) {
                Write-Verbose "Removing $dn from GrantSendOnBehalfTo list because user no longer has a mailbox"
                $sobo.Remove($dn)
                $dirty = $true
            }
        }

        if ($GrantSendOnBehalfTo) {
            # Grant SendOnBehalfOf rights to the owner, if appropriate
            if ($objUser.RecipientType -match 'MailUser' -or $objUser.RecipientType -match 'UserMailbox') {
                if ( !$sobo.Contains($objUser.DistinguishedName) ) {
                    $sobo.Add( $objUser.DistinguishedName )
                    $dirty = $true
                }
            } else {
                Write-Warning "Not granting user Send-On-Behalf-Of rights because $($objUser.SamAccountName) is a $($objUser.RecipientType)"
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
            Write-Verbose "Adding $objUser as a resource delegate on $resource"
            $resourceDelegates = (Get-CalendarProcessing -Identity $resource).ResourceDelegates
            if ( ! $resourceDelegates.Contains($objUser.DistinguishedName) ) {
                $resourceDelegates.Add($objUser.DistinguishedName)
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

        if ($EmailDelegate -and $objUser.RecipientTypeDetails -ne "User" ) {
            $Subject = "Information about Exchange resource `"$resource`""
            if ( $resourceType -eq 'SharedMailbox' ) {
                $Subject += " (Shared Mailbox)"
                $Template = $SharedMailboxTemplateEmail
            } elseif ( $resourceType -eq 'EquipmentMailbox' ) {
                $Subject += " (Equipment Resource)"
                $Template = $ResourceMailboxTemplateEmail
            } elseif ( $resourceType -eq 'RoomMailbox' ) {
                $Subject += " (Room Resource)"
                $Template = $ResourceMailboxTemplateEmail
            }

            $To = (Get-Recipient $objUser.Identity).PrimarySmtpAddress.ToString()

            $MessageBody = Resolve-TemplatedEmail `
                            -FilePath $Template `
                            -TemplateSubstitutions @{
                                "[[Subject]]"               = $Subject
                                "[[FirstName]]"             = $objUser.FirstName
                                "[[ResourceDisplayName]]"   = $resource.DisplayName
                                "[[ResourceEmailAddress]]"  = $resource.PrimarySmtpAddress.ToString()
                            }

            Send-MailMessage -To $To -From $From -Bcc $Bcc -Subject $Subject `
                             -BodyAsHtml -Body $MessageBody `
                             -SmtpServer $SmtpServer

            Write-Verbose "Sent message to $To for resource `"$resource`""
        }
    }
}
