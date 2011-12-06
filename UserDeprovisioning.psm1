################################################################################
#
# DESCRIPTION:  Deprovisions resources in Exchange for JMU
#
# Copyright (c) 2011 Seth Wright <wrightst@jmu.edu>
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

function Disable-ProvisionedMailbox {
    [CmdletBinding(SupportsShouldProcess=$true,
            ConfirmImpact="High")]

    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [Alias("Name")]
            # Specifies the user to be provisioned.
            $Identity,

            [Parameter(Mandatory=$true,
                ValueFromPipelineByPropertyName=$true)]
            [ValidateSet("Local", "Remote")]
            [string]
            # Should be either "Local" or "Remote"
            $MailboxLocation,

            [Parameter(Mandatory=$false)]
            [switch]
            # Specifies whether or not email notifications will be sent about the new mailbox.
            $SendEmailNotification=$true,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The domain controller to use for all operations.
            $DomainController,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # Where to create MailContact objects, if necessary.
            $MailContactOrganizationalUnit="ExchangeObjects/MailContacts"
        )

    # This section executes only once, before the pipeline.
    BEGIN {
        Write-Verbose "Performing initialization actions."

        if ([String]::IsNullOrEmpty($DomainController)) {
            $dc = [System.DirectoryServices.ActiveDirectory.Domain]::`
                    GetCurrentDomain().FindDomainController().Name

            if ($dc -eq $null) {
                Write-Error "Could not find a domain controller to use for the operation."
                continue
            }
        } else {
            $dc = $DomainController
        }

        Write-Verbose "Using Domain Controller $dc"
        Write-Verbose "Initialization complete."
    } # end 'BEGIN{}'

    # This section executes for each object in the pipeline.
    PROCESS {
        Write-Verbose "Beginning deprovisioning process for `"$Identity`""

        $resultObj = New-Object PSObject -Property @{
            Identity                    = $Identity
            MailboxLocation             = $MailboxLocation
            OriginalState               = $null
            EndingState                 = $null
            MailContactRemoved          = $false
            EmailSent                   = $false
            DeprovisioningSuccessful    = $false
            Error                       = $null
        }

        $User = $null
        try {
            Write-Verbose "Using Domain Controller $dc"
            $User = Get-User -Identity $Identity -DomainController $dc -ErrorAction Stop
        } catch {
            $err = "$Identity is not a valid user in Active Directory:  ($_)."
            Write-Error $err
            $resultObj.Error = $err
            return $resultObj
        }
        Write-Verbose "Found user $User in Active Directory"

        Write-Verbose "RecipientTypeDetails:  $($User.RecipientTypeDetails)"
        $username = $User.SamAccountName
        $resultObj.Identity = $username
        $resultObj.OriginalState = $User.RecipientTypeDetails
        $resultObj.EndingState = $User.RecipientTypeDetails

        # We don't process disabled users, at least not right now.
        if ($User.RecipientTypeDetails -eq 'DisabledUser') {
            $err = "User is disabled.  This may not work."
            Write-Warning $err
            $resultObj.Error = $err
        }

        # Find out if the user has a MailContact object.
        $contact = Get-MailContact -Identity "$($username)-mc" `
                        -DomainController $dc `
                        -ErrorAction SilentlyContinue
        if ($contact -ne $null) {
            Write-Verbose "Found MailContact object $($contact.Alias) for user $username"
        }

        if ($MailboxLocation -eq 'Local') {
            if ($User.RecipientTypeDetails -eq 'UserMailbox') {
                Write-Verbose "Requesting removal of local mailbox for UserMailbox $username"
            } else {
                # If 'Local' is specified, the user should have a local mailbox.
                $err = "$username does not have a local mailbox to remove."
                Write-Error $err
                $resultObj.Error = $err
                return $resultObj
            }
        } elseif ($MailboxLocation -eq 'Remote') {
            if ($User.RecipientTypeDetails -eq 'MailUser') {
                # If the user is a MailUser, we'll be disabling them as
                # such.
                Write-Verbose "Requesting removal of remote mailbox for MailUser $username"
            }

            if ($contact -ne $null) {
                Write-Verbose "Requesting removal of MailContact object for $username"
            } else {
                if ($User.RecipientTypeDetails -ne 'MailUser') { 
                    $err = "$username does not have a remote mailbox to remove."
                    Write-Error $err
                    $resultObj.Error = $err
                    return $resultObj
                }
            }
        }

        # TODO: Fix these comments.
        # After this, the following statements should be true:
        #  If the object is a User, there's nothing to do but cleanup a
        #  MailContact, if one exists.

        $desc = "Deprovision $MailboxLocation Mailbox for `"$username`""
        $caption = $desc
        $warning = "Are you sure you want to perform this action?`n"
        $warning += "This will remove the $($MailboxLocation.ToLower()) "
        $warning += "mailbox for user `"$username`"."
        if ($MailboxLocation -eq 'Remote') {
            $warning += "`n`nThis will NOT remove the remote mailbox, "
            $warning += "only the reference to the mailbox in the on-premise GAL."
        }

        if (!$PSCmdlet.ShouldProcess($desc, $warning, $caption)) {
            $resultObj.Error = "User cancelled the operation."
            return $resultObj
        }

        # Save some attributes that all user objects have.
        $savedAttributes = New-Object System.Collections.Hashtable
        $savedAttributes["DisplayName"] = $User.DisplayName
        $savedAttributes["SimpleDisplayName"] = $User.SimpleDisplayName

        # It's not a "saved attribute", per se, but it'll be handled in
        # the same way as a MailUser's ExternalEmailAddress attribute.
        if ([String]::IsNullOrEmpty($ExternalEmailAddress) -eq $false) {
            $savedAttributes["ExternalEmailAddress"] = $ExternalEmailAddress
        }

        # Save some attributes that tend to get blanked out.
        if ($User.RecipientTypeDetails -eq 'MailUser' -or
            $User.RecipientTypeDetails -eq 'UserMailbox') {

            try {
                $User = Get-Recipient $username -DomainController $dc -ErrorAction Stop
            } catch {
                $err = "Could not perform Get-Recipient on $username:  $_"
                Write-Error $err
                $resultObj.Error = $err
                return $resultObj
            }

            if ($User.RecipientTypeDetails -eq 'MailUser') {
                # Attributes that only MailUsers have.
                $savedAttributes["ExternalEmailAddress"] = $User.ExternalEmailAddress
                $savedAttributes["LegacyExchangeDN"] = (Get-MailUser $User.DistinguishedName).LegacyExchangeDN
            }

            # Attributes that both MailUsers and UserMailboxes have.
            $savedAttributes["CustomAttribute1"] = $User.CustomAttribute1
            $savedAttributes["CustomAttribute2"] = $User.CustomAttribute2
            $savedAttributes["CustomAttribute3"] = $User.CustomAttribute3
            $savedAttributes["CustomAttribute4"] = $User.CustomAttribute4
            $savedAttributes["CustomAttribute5"] = $User.CustomAttribute5
            $savedAttributes["CustomAttribute6"] = $User.CustomAttribute6
            $savedAttributes["CustomAttribute7"] = $User.CustomAttribute7
            $savedAttributes["CustomAttribute8"] = $User.CustomAttribute8
            $savedAttributes["CustomAttribute9"] = $User.CustomAttribute9
            $savedAttributes["CustomAttribute10"] = $User.CustomAttribute10
            $savedAttributes["CustomAttribute11"] = $User.CustomAttribute11
            $savedAttributes["CustomAttribute12"] = $User.CustomAttribute12
            $savedAttributes["CustomAttribute13"] = $User.CustomAttribute13
            $savedAttributes["CustomAttribute14"] = $User.CustomAttribute14
            $savedAttributes["CustomAttribute15"] = $User.CustomAttribute15
        }

        # Print out everything we know about the user so far (-Verbose)
        foreach ($key in ($savedAttributes.Keys | Sort)) {
            Write-Verbose "$($key):`t$($savedAttributes[$key])"
        }

        try {
            if ($MailboxLocation -eq 'Remote' -and $User.RecipientTypeDetails -eq 'MailUser') {
                Write-Verbose "Disabling $username as a MailUser"
                Disable-MailUser `
                    -Identity $username `
                    -DomainController $dc `
                    -Confirm:$false `
                    -ErrorAction Stop
            } elseif ($MailboxLocation -eq 'Local' -and $User.RecipientTypeDetails -eq 'UserMailbox') {
                Write-Verbose "Disabling $username as a UserMailbox"
                Disable-Mailbox `
                    -Identity $username `
                    -DomainController $dc `
                    -Confirm:$false `
                    -ErrorAction Stop
            }
        } catch {
            $err = "Could not run Disable-Mailbox/Disable-MailUser for user $username ($_)."
            Write-Error $err
            $resultObj.Error = $err
            return $resultObj
        }

        # Object should now just be a "User".
        $User = Get-User -Identity $username -DomainController $dc -ErrorAction Stop
        $resultObj.EndingState = $User.RecipientTypeDetails

        if ($contact -ne $null) {
            # Save some relevant attributes.
            $mcExternalEmailAddress = $contact.ExternalEmailAddress
            $mcEmailAddresses = $contact.EmailAddresses

            Write-Verbose "MailContact's ExternalEmailAddress: $mcExternalEmailAddress"
            Write-Verbose "MailContact's EmailAddresses: $mcEmailAddresses"

            try {
                Write-Verbose "Removing MailContact object $($contact.Alias)"
                Remove-MailContact `
                    -Identity "$($username)-mc" `
                    -DomainController $dc `
                    -Confirm:$Confirm `
                    -ErrorAction Stop
            } catch {
                $err =  "Could not remove contact for $username.  The error was:  $_"
                Write-Error $err
                $resultObj.Error = $_
                return $resultObj
            }

            $resultObj.MailContactRemoved = $true

            if ($User.RecipientTypeDetails -eq 'User' -and $MailboxLocation -eq 'Local') {
                # ...which it should...
                Write-Verbose "Enabling $username as a MailUser"

                try {
                    $User = Enable-MailUser `
                                -Identity $username `
                                -ExternalEmailAddress $mcExternalEmailAddress `
                                -DomainController $dc `
                                -ErrorAction Stop
                } catch {
                    $err = "An error occurred while running Enable-MailUser:  $_"
                    Write-Error $err
                    $resultObj.Error = $err
                    return $resultObj
                }

                # TODO:  Handle removal of @jmu.edu addresses.
                $emailAddresses = $User.EmailAddresses
                foreach ($address in $contact.EmailAddresses) {
                    if ($address.PrefixString -match 'X500') {
                        if ($address.AddressString -eq $User.LegacyExchangeDN) {
                            Write-Verbose "LegacyExchangeDN is already correct"
                        } else {
                            $emailAddresses += $address
                        }
                    }
                }

                $resultObj.EndingState = $User.RecipientTypeDetails
            }
        }

        # Reapply any saved attributes.
        if ($User.RecipientTypeDetails -eq 'MailUser') {
            $cmd = "Set-MailUser -Identity $Identity -DomainController $dc -ErrorAction Stop "
            $cmd += " -EmailAddresses `$emailAddresses "

            foreach ($key in $savedAttributes.Keys) {
                if ($key -eq 'LegacyExchangeDN' -or
                    $key -eq 'ExternalEmailAddress' -or
                    [String]::IsNullOrEmpty($savedAttributes[$key])) {
                    continue
                }
                $cmd += "-$($key) `"$($savedAttributes[$key])`" "
            }

            Write-Verbose "Reapplying saved attributes"

            try {
                Write-Debug "Executing command `"$cmd`""
                Invoke-Expression $cmd
            } catch {
                $err =  "An error occurred while reapplying saved attributes.  The error was: $_"
                Write-Error $err
                $resultObj.Error = $err
                return $resultObj
            }
        }

        $resultObj.DeprovisioningSuccessful = $true
        $resultObj.Error = "$MailboxLocation mailbox deprovisioned."
        return $resultObj
    } # end 'PROCESS{}'

    # This section executes only once, after the pipeline.
    END {
        Write-Verbose "Cleaning up"
    } # end 'END{}'
}
