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

function Remove-ProvisionedMailbox {
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

            [Parameter(Mandatory=$false,
                ParameterSetName="MailUser",
                ValueFromPipelineByPropertyName=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The external address (targetAddress) to assign to the MailUser.
            $ExternalEmailAddress,

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
#        if ($User.RecipientTypeDetails -eq 'DisabledUser') {
#            $err = "User is disabled in Active Directory.  Not performing any operations."
#            Write-Error $err
#            $resultObj.Error = $err
#            return $resultObj
#        }

        # If 'Local' is specified, the user should have a local mailbox.
        if ($MailboxLocation -eq 'Local' -and $User.RecipientTypeDetails -ne 'UserMailbox') {
            $resultObj.DeprovisioningSuccessful = $true
            $err = "$username does not have a local mailbox to remove."
            Write-Error $err
            $resultObj.Error = $err
            return $resultObj
        }

        # If 'Remote' is specified, the user should have a remote mailbox.
        if ($MailboxLocation -eq 'Remote' -and $User.RecipientTypeDetails -ne 'MailUser') {
            $resultObj.DeprovisioningSuccessful = $true
            $err = "$username does not have a remote mailbox to remove."
            Write-Error $err
            $resultObj.Error = $err
            return $resultObj
        }


        # TODO: Fix these comments.
        # After this, the following statements should be true:
        #  If the object is a User, either "Local" or "Remote" can be specified.
        #  If the object is a MailUser, then "Local" was specified.
        #  If the object is a UserMailbox, then "Remote" was specified.

        $desc = "Deprovision $MailboxLocation Mailbox for `"$username`""
        $caption = $desc
        $warning = "Are you sure you want to perform this action?`n"
        $warning += "This will remove the $($MailboxLocation.ToLower()) "
        $warning += "mailbox for user `"$username`"."

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
            if ($User.RecipientTypeDetails -eq 'MailUser') {
                Write-Verbose "Disabling $username as a MailUser"
                Disable-MailUser `
                    -Identity $username `
                    -DomainController $dc `
                    -Confirm:$false `
                    -ErrorAction Stop
            } elseif ($User.RecipientTypeDetails -eq 'UserMailbox') {
                Write-Verbose "Disabling $username as a UserMailbox"
                Disable-Mailbox `
                    -Identity $username `
                    -DomainController $dc `
                    -Confirm:$false `
                    -ErrorAction Stop
            }
        } catch {
            $err = "Could not run Disable-Mailbox/Disable-Mailuser for user $username ($_)."
            Write-Error $err
            $resultObj.Error = $err
            return $resultObj
        }

        # Object should now just be a "User".
        $User = Get-User -Identity $username -DomainController $dc -ErrorAction Stop
        $resultObj.EndingState = $User.RecipientTypeDetails

        # If a mail contact object exists for the user, it should be deleted
        # the user should become a MailUser with that information.
        $contact = Get-MailContact -Identity "$($username)-mc" `
                        -DomainController $dc `
                        -ErrorAction SilentlyContinue

        if ($contact -eq $null) {
            Write-Verbose "No MailContact exists for $username"
        } else {
            # Save some relevant attributes.
            $mcExternalEmailAddress = $contact.ExternalEmailAddress
            $mcEmailAddresses = $contact.EmailAddresses

            try {
                Remove-MailContact `
                    -Identity "$($username)-mc)"
                    -DomainController $dc `
                    -ErrorAction Stop
            } catch {
                $err =  "Could not remote contact for $username.  The error was:  $_"
                Write-Error $err
                $resultObj.Error = $_
                return $resultObj
            }
        }

        $resultObj.MailContactRemoved = $true

        if ([String]::IsNullOrEmpty($savedAttributes["LegacyExchangeDN"]) -eq $false) {
            # NOTE:  The "X" is lower-case on purpose in order to make 
            # it a secondary address.
            # Not that, you know, it works.  But it should.  As soon as
            # I add the address to the collection, it forces it to be a
            # primary address, and it won't change it back.
            $addr = "x500:" + $savedAttributes["LegacyExchangeDN"]

            # Only add the X500 address if it doesn't already exist.
            if ($contact.EmailAddresses.Contains($addr) -eq $false) {
                Write-Verbose "Adding LegacyExchangeDN to new MailContact as an X500 address"
                $contact.EmailAddresses.Add($addr)

                try {
                    Write-Verbose "Attempting to set new EmailAddresses for user"
                    Set-MailContact -Identity $contact.Identity `
                        -EmailAddresses $contact.EmailAddresses `
                        -DomainController $dc `
                        -ErrorAction Stop
                } catch {
                    $w = "An error occurred while adding " + $savedAttributes["LegacyExchangeDN"]
                    $w += " as an X500 address to the MailContact for $username.  "
                    $w += "You will need to add this manually.  The error was:  $_"
                    Write-Warning $w
                    $resultObj.Error = $w
                }
            } else {
                Write-Verbose "LegacyExchangeDN already exists as an X500 address"
            }

            if ($contact -ne $null) {
                $contact = Get-MailContact -Identity $contact.Identity -DomainController $dc
            }
        } else {
            Write-Verbose "Did not find a LegacyExchangeDN for user"
        }

        # At this point, the following should be true:
        # * If "Local" was specified--no matter what type the object was as
        #   the start--then the object will be enabled as a UserMailbox.
        #
        # * If "Remote" was specified and the object is a UserMailbox, a
        #   MailContact has been created and that's all that needs to happen.
        #
        # * If "Remote" was specified and the object is a User, the object will
        #   be enabled as a MailUser.

        if ($MailboxLocation -eq 'Local') {
            # Enable as UserMailbox.
            Write-Verbose "Enabling $username as a UserMailbox"

            try {
                # TODO:  Remove the JMU-specific bit about the ManagedFolderMailboxPolicy.
                $User = Enable-Mailbox `
                            -Identity $username `
                            -ManagedFolderMailboxPolicy "Default Managed Folder Policy" `
                            -ManagedFolderMailboxPolicyAllowed:$true `
                            -DomainController $dc `
                            -ErrorAction Stop
            } catch {
                $err =  "An error occurred while running Enable-Mailbox.  The error was: $_"
                Write-Error $err
                $resultObj.Error = $err
                return $resultObj
            }
        } elseif ($MailboxLocation -eq 'Remote' -and $User.RecipientTypeDetails -eq 'User') {
            # Enable as MailUser.
            Write-Verbose "Enabling $username as a MailUser"

            try {
                $User = Enable-MailUser `
                            -Identity $username `
                            -ExternalEmailAddress $savedAttributes["ExternalEmailAddress"] `
                            -DomainController $dc `
                            -ErrorAction Stop
            } catch {
                $err = "An error occurred while running Enable-MailUser:  $_"
                Write-Error $err
                $resultObj.Error = $err
                return $resultObj
            }
        }

        $resultObj.EndingState = $User.RecipientTypeDetails

        # Reapply any saved attributes.
        if ($User.RecipientTypeDetails -eq 'MailUser') {
            $cmd = "Set-MailUser "
        } elseif ($User.RecipientTypeDetails -eq 'UserMailbox') {
            $cmd = "Set-Mailbox "
        }
        $cmd += "-Identity $Identity -DomainController $dc "

        foreach ($key in $savedAttributes.Keys) {
            if ($key -eq 'LegacyExchangeDN' -or
                $key -eq 'ExternalEmailAddress' -or
                [String]::IsNullOrEmpty($savedAttributes[$key])) {
                continue
            }
            $cmd += "-$($key) `"$($savedAttributes[$key])`" "
        }
        $cmd += "-ErrorAction Stop"

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

        # Send emails where appropriate.
        if ($SendEmailNotification -eq $true) {
            $welcomeEmail = $null
            $notifyEmail = $null

            if ($MailboxLocation -eq "Local") {
                $welcomeEmail = $User.PrimarySmtpAddress
                if ($contact -ne $null) {
                    $notifyEmail = $contact.ExternalEmailAddress
                }
            } elseif ($MailboxLocation -eq "Remote") {
                if ($contact -eq $null) {
                    $welcomeEmail = $User.ExternalEmailAddress
                } else {
                    $welcomeEmail = $contact.ExternalEmailAddress
                    $notifyEmail = $User.PrimarySmtpAddress
                }
            }

            if ($welcomeEmail -ne $null) {
                Write-Verbose "Sending Welcome email to $welcomeEmail"
            }
            if ($notifyEmail -ne $null) {
                Write-Verbose "Sending Notify email to $notifyEmail"
            }
        } else {
            Write-Verbose "Not sending emails"
        }

        $resultObj.DeprovisioningSuccessful = $true
        $resultObj.Error = "$MailboxLocation mailbox provisioned."
        return $resultObj
    } # end 'PROCESS{}'

    # This section executes only once, after the pipeline.
    END {
        Write-Verbose "Cleaning up"
    } # end 'END{}'
}
