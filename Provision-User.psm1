################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  Provisions users in Exchange.
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

function Provision-User {
    [CmdletBinding(SupportsShouldProcess=$true,
            ConfirmImpact="High")]

    param ( 
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [Alias("Identity")]
            [Alias("Name")]
            # Specifies the user to be provisioned.
            $User,

            [Parameter(Mandatory=$false)]
            # Whether to force-create a mailbox for a user, even if they would
            # not normally be a candidate for a mailbox
            [switch]
            $Force,

            [Parameter(Mandatory=$true,
                ParameterSetName="Mailbox")]
            [switch]
            # User should be provisioned as a UserMailbox.
            $Mailbox,

            [Parameter(Mandatory=$true,
                ParameterSetName="MailUser")]
            [switch]
            # User should be provisioned as a MailUser.
            $MailUser,

            [Parameter(Mandatory=$true,
                ParameterSetName="MailUser")]
            [string]
            # The external address (targetAddress) to assign to the MailUser.
            $ExternalEmailAddress,

            [Parameter(Mandatory=$false)]
            # Extra email addresses to add to the recipient.
            $EmailAddresses,

            [Parameter(Mandatory=$false)]
            [switch]
            # Whether this function is being called in an automated fashion.
            $Automated,

            [Parameter(Mandatory=$false)]
            [string]
            # The file in which to write out information about users who were
            # not successfully provisioned, if the Automated parameter is
            # specified.
            $RetryableErrorsFilePath,

            [Parameter(Mandatory=$false)]
            $DomainController
        )

# This section executes only once, before the pipeline.
    BEGIN {
        Write-Verbose "Performing initialization actions."

        if ([String]::IsNullOrEmpty($DomainController)) {
            $dc = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController().Name

            if ($dc -eq $null) {
                Write-Error "Could not find a domain controller to use for the operation."
                return
            }
        } else {
            $dc = $DomainController
        }
        Write-Verbose "Using Domain Controller $dc"

        $defaultStart = Get-Date

        if ($Automated -eq $true) {
            if ([String]::IsNullOrEmpty($RetryableErrorsFilePath)) {
                $errorsFile = [System.IO.Path]::Combine($(Get-Location), "provisioning_errors_$($defaultStart.Ticks).csv")
            } else {
                $errorsFile = $RetryableErrorsFilePath
            }

            Write-Verbose "Errors File:  $errorsFile"
        }

        Write-Verbose "Start time:  $defaultStart"

        Write-Verbose "Initialization complete."
    } # end 'BEGIN{}'

# This section executes for each object in the pipeline.
    PROCESS {
        Write-Verbose "Beginning provisioning process for `"$User`""

        if ([String]::IsNullOrEmpty($User)) { 
            Write-Error "User was null."
            return
        }

        if ($User.User -eq $null) {
            $userName = $User
            $start = $defaultStart
        } else {
            $userName = $User.User
            $start = $User.Date
        }

        if ($EmailAddresses -eq $null) {
            $EmailAddresses = New-Object System.Collections.ArrayList
        }

        Write-Verbose "Finding user in Active Directory"
        $objUser = Get-User $userName -ErrorAction SilentlyContinue

        if ($objUser -eq $null) {
            Write-Error "$userName is not a valid user in Active Directory."
            if ($Automated) {
                Out-File -NoClobber -Append -FilePath $errorsFile -InputObject "$userName,$start,No AD Account"
            }
            return
        }

        $userName = $objUser.SamAccountName
        Write-Verbose "sAmAccountName:  $userName"

        # Save these off because Exchange blanks them out...
        $displayName = $objUser.DisplayName
        $displayNamePrintable = $objUser.SimpleDisplayName
        Write-Verbose "displayName:  $displayName"
        Write-Verbose "displayNamePrintable: $displayNamePrintable"

        # Perform some sanity checks and save off some data before getting into the
        # actual provisioning process.
        Write-Verbose "$userName is a $($objUser.RecipientTypeDetails)"
        switch ($objUser.RecipientTypeDetails) {
            'User' { 
                # Nothing special about the object.
                break
            }
            'MailUser' {
                if ($MailUser -eq $true) {
                    # If the object is a MailUser, and we are trying to enable it
                    # as a MailUser, error out.
                    Write-Error "$userName is already a MailUser."
                    return
                }

                # If we're "upgrading" someone, save off any extra attributes that
                # only Exchange recipients have
                Write-Verbose "Saving Exchange-specific attributes in order to reapply them later"
                $recip = Get-MailUser $userName
                $externalEmailAddress = $recip.ExternalEmailAddress
                $legacyExchangeDn = $recip.LegacyExchangeDN

                $customAttribute1  = $recip.CustomAttribute1
                $customAttribute2  = $recip.CustomAttribute2
                $customAttribute3  = $recip.CustomAttribute3
                $customAttribute4  = $recip.CustomAttribute4
                $customAttribute5  = $recip.CustomAttribute5
                $customAttribute6  = $recip.CustomAttribute6
                $customAttribute7  = $recip.CustomAttribute7
                $customAttribute8  = $recip.CustomAttribute8
                $customAttribute9  = $recip.CustomAttribute9
                $customAttribute10 = $recip.CustomAttribute10
                $customAttribute11 = $recip.CustomAttribute11
                $customAttribute12 = $recip.CustomAttribute12
                $customAttribute13 = $recip.CustomAttribute13
                $customAttribute14 = $recip.CustomAttribute14
                $customAttribute15 = $recip.CustomAttribute15

                $currentAddrs = $recip.EmailAddresses.Clone()
                # Save any extra addresses that aren't the same as the
                # targetAddress.
                foreach ($addr in $currentAddrs) {
                    if ($addr -notmatch $recip.ExternalEmailAddress) { 
                        $EmailAddresses.Add($addr.ToSecondary()) | Out-Null
                    }
                }

                break
            }
            'UserMailbox' {
                if ($Mailbox -eq $true) {
                    # If the object is a UserMailbox, and we are trying to enable
                    # it as a UserMailbox, error out.
                    Write-Error "$userName already has a mailbox."
                    return
                } elseif ($MailUser -eq $true) {
                    # If the object is a UserMailbox, and we are tryint to enable
                    # it as a MailUser, refuse to downgrade and error out.
                    Write-Error "$userName is a Mailbox, refusing to enable as MailUser instead"
                    return
                }
                break
            }
            'DisabledUser' {
                if ($Mailbox -eq $true) {
                    Write-Error "$userName is disabled, refusing to create mailbox."
                    if ($Automated) {
                        Out-File -Append -FilePath $errorsFile -InputObject "$userName,$start,Disabled AD Account"
                    }
                    return
                }
                break
            }
            default {
                # If the object is anything else, we don't know how to deal with
                # it.
                Write-Error "$userName is a $($objUser.RecipientTypeDetails) object, refusing to provision."
                return
            }
        }

        
        if ($Mailbox -eq $true) {
            Write-Verbose "DistinguishedName:  $($objUser.DistinguishedName)"
            # Don't auto-create mailboxes for users in the Students OU
            if ($objUser.DistinguishedName -match 'Student') {
                Write-Debug "Force is $Force"
                if ($Force -eq $false) {
                    Write-Error "$userName is listed as a student, refusing to create mailbox."
                    # User chose to cancel the operation.
                    if ($Automated) {
                        Out-File -Append -FilePath $errorsFile -InputObject "$userName,$start,Student"
                    }
                    return
                } else {
                    Write-Warning "$userName is listed as a student; confirmation is required in order to create a mailbox for the user."
                    if (!$PSCmdlet.ShouldProcess($userName, "create mailbox")) {
                        Write-Verbose "Operation cancelled."
                        # User chose to cancel the operation.
                        if ($Automated) {
                            Out-File -Append -FilePath $errorsFile -InputObject "$userName,$start,Student"
                        }
                        return
                    }
                }
            }

            # If the user is a MailUser already, since we're trying to enable them
            # as a UserMailbox we need to disable them as a MailUser first.
            if ($objUser.RecipientTypeDetails -match 'MailUser') {
                Write-Verbose "Disabling $userName as a MailUser in order to enable it as a UserMailbox"
                $error.Clear()
                Deprovision-User -User $userName -Confirm:$false -DomainController $dc -Verbose $Verbose

                if (![String]::IsNullOrEmpty($error[0])) {
                    Write-Error $error[0]
                    if ($Automated) {
                        Out-File -Append -FilePath $errorsFile -InputObject "$userName,$start,Deprovisioning Error"
                    }
                    return
                }
            }

            # Enable the mailbox
            Write-Verbose "Enabling $userName as a UserMailbox"
            $Error.Clear()
            Enable-Mailbox -Identity $userName `
                -ManagedFolderMailboxPolicy "Default Managed Folder Policy" `
                -ManagedFolderMailboxPolicyAllowed:$true `
                -DomainController $dc -ErrorAction SilentlyContinue

            if ($Error[0] -ne $null) {
                Write-Error $Error[0]
                if ($Automated) {
                    Out-File -Append -FilePath $errorsFile -InputObject "$userName,$start,$($Error[0].ToString())"
                }
                return
            } 

            $error.Clear()
            # Set various attributes now that Exchange has helpfully removed them.
            Write-Verbose "Resetting displayName and displayNamePrintable attributes to saved values"
            Set-User -Identity $userName -DisplayName $displayName `
                -SimpleDisplayName $displayNamePrintable `
                -DomainController $dc `
                -ErrorAction SilentlyContinue

            # ...and if the object was a MailUser before, reset all of the Custom
            # Attributes as well.
            if ($objUser.RecipientTypeDetails -eq "MailUser") {
                Write-Verbose "Resetting CustomAttributes (extensionAttribute*) 1 - 15 to saved values"
                Set-Mailbox -Identity $userName `
                    -CustomAttribute1 $customAttribute1 `
                    -CustomAttribute2 $customAttribute2 `
                    -CustomAttribute3 $customAttribute3 `
                    -CustomAttribute4 $customAttribute4 `
                    -CustomAttribute5 $customAttribute5 `
                    -CustomAttribute6 $customAttribute6 `
                    -CustomAttribute7 $customAttribute7 `
                    -CustomAttribute8 $customAttribute8 `
                    -CustomAttribute9 $customAttribute9 `
                    -CustomAttribute10 $customAttribute10 `
                    -CustomAttribute11 $customAttribute11 `
                    -CustomAttribute12 $customAttribute12 `
                    -CustomAttribute13 $customAttribute13 `
                    -CustomAttribute14 $customAttribute14 `
                    -CustomAttribute15 $customAttribute15 `
                    -DomainController $dc `
                    -ErrorAction SilentlyContinue
            }

            if ($EmailAddresses -eq $null) {
                Write-Verbose "Not processing EmailAddresses; parameter was null."
            } else {
                Write-Verbose "Adding email addresses to $($userName)'s EmailAddresses collection"

                $error.Clear()
                $addrs = (Get-Recipient -Identity $userName).EmailAddresses
                if (![String]::IsNullOrEmpty($error[0])) {
                    Write-Error $error[0]
                }

                foreach ($addr in $addrs) { 
                    Write-Debug "Working on address $addr"
                    if (!$EmailAddresses.Contains($addr)) {
                        Write-Verbose "Adding $addr to the collection"
                        $EmailAddresses.Add($addr) | Out-Null
                    }
                }

                Write-Verbose "Setting Email Addresses for $userName"
                $error.Clear()
                Set-Mailbox -Identity $userName `
                    -EmailAddressPolicyEnabled:$false `
                    -EmailAddresses $EmailAddresses `
                    -DomainController $dc
                if (![String]::IsNullOrEmpty($error[0])) {
                    Write-Error $error[0]
                }

                Write-Debug "Reapplying Email Address Policy"
                $error.Clear()
                Set-Mailbox -Identity $userName `
                    -EmailAddressPolicyEnabled:$true `
                    -DomainController $dc
                if (![String]::IsNullOrEmpty($error[0])) {
                    Write-Error $error[0]
                }
            }
        } elseif ($MailUser -eq $true) {
            # The user should be enabled as a MailUser instead of a Mailbox.
            Write-Verbose "Enabling $userName as a MailUser"
            $Error.Clear()
            Enable-MailUser -Identity $userName `
                -ExternalEmailAddress $ExternalEmailAddress `
                -DomainController $dc

            if (![String]::IsNullOrEmpty($error[0])) {
                Write-Error $error[0]
                if ($Automated) {
                    Out-File -Append -FilePath $errorsFile -InputObject "$userName,$start,$($Error[0].ToString())"
                }
                return
            } 

            if ($EmailAddresses -eq $null) {
                Write-Verbose "Not processing EmailAddresses; parameter was null."
            } else {
                Write-Verbose "Explicitly adding email addresses to $($userName)'s EmailAddresses collection"

                $error.Clear()
                $addrs = (Get-Recipient -Identity $userName).EmailAddresses
                if (![String]::IsNullOrEmpty($error[0])) {
                    Write-Error $error[0]
                }

                foreach ($addr in $addrs) { 
                    Write-Debug "Working on address $addr"
                    if (!$EmailAddresses.Contains($addr)) {
                        Write-Verbose "Adding $addr to the collection"
                        $EmailAddresses.Add($addr) | Out-Null
                    }
                }

                Write-Verbose "Setting Email Addresses for $userName"
                $error.Clear()
                Set-MailUser -Identity $userName `
                    -EmailAddressPolicyEnabled:$false `
                    -EmailAddresses $EmailAddresses `
                    -DomainController $dc
                if (![String]::IsNullOrEmpty($error[0])) {
                    Write-Error $error[0]
                }

                Write-Debug "Reapplying Email Address Policy"
                $error.Clear()
                Set-MailUser -Identity $userName `
                    -EmailAddressPolicyEnabled:$true `
                    -DomainController $dc
                if (![String]::IsNullOrEmpty($error[0])) {
                    Write-Error $error[0]
                }
            }
        }

        Write-Verbose "Ending provisioning process for `"$User`""
    } # end 'PROCESS{}'

# This section executes only once, after the pipeline.
    END {
        Write-Verbose "Cleaning up"
    } # end 'END{}'
}
