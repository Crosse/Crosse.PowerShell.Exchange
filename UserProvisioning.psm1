################################################################################
#
# DESCRIPTION:  Provisions resources in Exchange for JMU
#
# Copyright (c) 2009-2011 Seth Wright <wrightst@jmu.edu>
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

################################################################################
<#
    .SYNOPSIS
    Provisions a user with a local or remote mailbox.

    .DESCRIPTION
    Provisions a user with a local or remote mailbox.  This cmdlet will not actually
    create a mailbox in a remote Exchange instance (like Office365 or Live@edu),
    but will set up the user account in the on-premise Exchange correctly.

    .INPUTS
    System.String.  The Identity to provision.

    .OUTPUTS
    A PSObject with the results of the provisioning process.
#>
################################################################################
function Add-ProvisionedMailbox {
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
            # Should be either "Local" or "Remote".
            $MailboxLocation,

            [Parameter(Mandatory=$false,
                ValueFromPipelineByPropertyName=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The external address (targetAddress) to assign to the MailUser.
            $ExternalEmailAddress,

            [Parameter(Mandatory=$false)]
            [switch]
            # Specifies whether or not email notifications will be sent about the new mailbox.
            $SendEmailNotification=$false,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [ValidateScript({ (Test-Path $_) })]
            [string]
            # The path to a file containing the template used to send the "welcome" email to
            # a user who receives a local mailbox.
            $LocalWelcomeEmailTemplate = (Join-Path $PSScriptRoot "LocalWelcomeEmailTemplate.html"),

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [ValidateScript({ (Test-Path $_) })]
            [string]
            # The path to a file containing the template used to send the "welcome" email to
            # a user who receives a remote mailbox.
            $RemoteWelcomeEmailTemplate = (Join-Path $PSScriptRoot "RemoteWelcomeEmailTemplate.html"),

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [ValidateScript({ (Test-Path $_) })]
            [string]
            # The path to a file containing the template used to send the "notification" email to
            # a user who receives a local mailbox.
            $LocalNotificationEmailTemplate = (Join-Path $PSScriptRoot "LocalNotificationEmailTemplate.html"),

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [ValidateScript({ (Test-Path $_) })]
            [string]
            # The path to a file containing the template used to send the "notification" email to
            # a user who receives a remote mailbox.
            $RemoteNotificationEmailTemplate = (Join-Path $PSScriptRoot "RemoteNotificationEmailTemplate.html"),

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [ValidateScript({ (Test-Path $_) })]
            [string]
            # The directory in which to place a CSV file containing email welcome/notification recipients.
            $EmailNotificationPath = (Join-Path $PWD "EmailNotifications"),

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The domain controller to use for all operations.
            $DomainController,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # Where to create MailContact objects, if necessary.
            $MailContactOrganizationalUnit="ExchangeObjects/DukesContacts"
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

        if ($SendEmailNotification -eq $true) {
            if ([String]::IsNullOrEmpty($LocalWelcomeEmailTemplate) -or
                    [String]::IsNullOrEmpty($LocalNotificationEmailTemplate) -or
                    [String]::IsNullOrEmpty($RemoteWelcomeEmailTemplate) -or
                    [String]::IsNullOrEmpty($RemoteNotificationEmailTemplate)) {
                Write-Error "Not all parameters found to support sending email."
                continue
            }
        }

        if ([String]::IsNullOrEmpty($EmailNotificationPath) -eq $false) {
            $emailNotificationFile = Join-Path $EmailNotificationPath "emailnotifications_$(Get-Date -Format yyyy-MM-dd_HH-mm-ss-fff).csv"
            $emailNotifications = New-Object System.Collections.ArrayList
            Write-Verbose "Will record email notifies in file $emailNotificationFile"
        }

        Write-Verbose "Using Domain Controller $dc"
        Write-Verbose "Initialization complete."
    } # end 'BEGIN{}'

    # This section executes for each object in the pipeline.
    PROCESS {
        Write-Verbose "Beginning provisioning process for `"$Identity`""

        $resultObj = New-Object PSObject -Property @{
            Identity                = $Identity
            MailboxLocation         = $MailboxLocation
            OriginalState           = $null
            EndingState             = $null
            MailContactCreated      = $false
            EmailSent               = $false
            ProvisioningSuccessful  = $false
            Error                   = $null
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
            $err = "User is disabled in Active Directory.  Not performing any operations."
            Write-Error $err
            $resultObj.Error = $err
            return $resultObj
        }

        # There is no sense in processing a UserMailbox if "Local"
        # was specified, since it already has a mailbox.
        if ($User.RecipientTypeDetails -eq 'UserMailbox' -and $MailboxLocation -eq 'Local') {
            $resultObj.ProvisioningSuccessful = $true
            $err = "$username already has a local mailbox."
            Write-Error $err
            $resultObj.Error = $err
            return $resultObj
        }

        # Another condition that doesn't make sense:  a MailUser getting a
        # "Remote" mailbox.
        if ($User.RecipientTypeDetails -eq 'MailUser' -and $MailboxLocation -eq 'Remote') {
            $resultObj.ProvisioningSuccessful = $true
            $err = "$username already has a remote mailbox."
            Write-Error $err
            $resultObj.Error = $err
            return $resultObj
        }

        # If "Remote" was specified, then an ExternalEmailAddress must also be present.
        if ($MailboxLocation -eq 'Remote' -and [String]::IsNullOrEmpty($ExternalEmailAddress)) {
            $err = "An ExternalEmailAddress must be provided for the Remote mailbox type."
            Write-Error $err
            $resultObj.Error = $err
            return $resultObj
        }


        # After this, the following statements should be true:
        #  If the object is a User, either "Local" or "Remote" can be specified.
        #  If the object is a MailUser, then "Local" was specified.
        #  If the object is a UserMailbox, then "Remote" was specified.

        $desc = "Provision $MailboxLocation Mailbox for `"$username`""
        $caption = $desc
        $warning = "Are you sure you want to perform this action?`n"
        $warning += "This will give the user `"$username`" a "
        $warning += "$($MailboxLocation.ToLower()) mailbox."

        if (!$PSCmdlet.ShouldProcess($desc, $warning, $caption)) {
            $resultObj.Error = "User cancelled the operation."
            return $resultObj
        }

        # Save some attributes that all user objects have.
        $savedAttributes = New-Object System.Collections.Hashtable
        $savedAttributes["DisplayName"] = $User.DisplayName
        $savedAttributes["SimpleDisplayName"] = $User.SimpleDisplayName
        $savedAttributes["FirstName"] = $User.FirstName

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
                $err = "Could not perform Get-Recipient on ${username}:  $_"
                Write-Error $err
                $resultObj.Error = $err
                return $resultObj
            }

            if ($User.RecipientTypeDetails -eq 'MailUser') {
                # Attributes that only MailUsers have.
                $savedAttributes["ExternalEmailAddress"] = $User.ExternalEmailAddress
                $mailUser = Get-MailUser -Identity $User.DistinguishedName `
                                         -DomainController $dc
                $savedAttributes["LegacyExchangeDN"] = $mailUser.LegacyExchangeDN
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
        }

        # This will be the "last touched time" attribute and should
        # be set for object.
        $savedAttributes["CustomAttribute15"] = "LastProvisioned: $(Get-Date -Format u)"

        # Print out everything we know about the user so far (-Verbose)
        foreach ($key in ($savedAttributes.Keys | Sort-Object)) {
            Write-Verbose "$($key):`t$($savedAttributes[$key])"
        }

        if ($User.RecipientTypeDetails -eq 'MailUser') {
            # MailUser --> UserMailbox:  user must first be disabled as a
            # MailUser, then it can be enabled as a mailbox.
            Write-Verbose "Disabling $username as a MailUser"

            try {
                Disable-MailUser `
                    -Identity $username `
                    -DomainController $dc `
                    -Confirm:$false `
                    -ErrorAction Stop
            } catch {
                $err = "Could not run Disable-MailUser for user $username ($_)."
                Write-Error $err
                $resultObj.Error = $err
                return $resultObj
            }

            # Object should now just be a "User".
            $User = Get-User -Identity $username -DomainController $dc -ErrorAction Stop
            $resultObj.EndingState = $User.RecipientTypeDetails

            # Since this user will become a UserMailbox, a MailContact needs
            # to be created to preserve their current routing information.
            $createContact = $true
        } elseif ($User.RecipientTypeDetails -eq 'UserMailbox') {
            # Remote mailbox; requires a contact to be created.
            $createContact = $true
        }

        # Create a MailContact for the user using the saved details so that
        # both addresses get listed in the GAL, etc.
        if ($createContact -eq $true)
        {
            Write-Verbose "Creating MailContact for $username"

            $contact = Get-MailContact -Identity "$($username)-mc" `
                            -DomainController $dc `
                            -ErrorAction SilentlyContinue

            if ($contact -eq $null) {
                try {
                    # TODO:  Remove the JMU-specific "(Dukes)" bit in the DisplayName.
                    # I don't know what to replace it with, though.  "(External)"?
                    $contact = New-MailContact `
                                    -Name "$($username)-mc" `
                                    -Alias "$($username)-mc" `
                                    -DisplayName "$($User.DisplayName) (Dukes)" `
                                    -FirstName "$($User.FirstName)" `
                                    -LastName "$($User.LastName)" `
                                    -ExternalEmailAddress $savedAttributes["ExternalEmailAddress"] `
                                    -OrganizationalUnit $MailContactOrganizationalUnit `
                                    -DomainController $dc `
                                    -ErrorAction Stop
                } catch {
                    $err =  "Could not create contact for $username.  The error was:  $_"
                    Write-Error $err
                    $resultObj.Error = $_
                    return $resultObj
                }
            } else {
                Write-Warning "MailContact already exists for $username"
                try {
                    # TODO:  Remove the JMU-specific "(Dukes)" bit in the DisplayName.
                    # I don't know what to replace it with, though.  "(External)"?
                    # Also, -FirstName and -LastName don't exist for Set-MailContact,
                    # so use Set-Contact to update those.
                    Write-Verbose "Updating MailContact attributes"
                    Set-MailContact -Identity "$($username)-mc" `
                        -Name "$($username)-mc" `
                        -Alias "$($username)-mc" `
                        -DisplayName "$($User.DisplayName) (Dukes)" `
                        -ExternalEmailAddress $savedAttributes["ExternalEmailAddress"] `
                        -DomainController $dc `
                        -ErrorAction Stop
                    Set-Contact -Identity "$($username)-mc" `
                        -FirstName "$($User.FirstName)" `
                        -LastName "$($User.LastName)" `
                        -DomainController $dc `
                        -ErrorAction Stop
                } catch {
                    $err =  "Could not modify existing contact for $username.  The error was:  $_"
                    Write-Error $err
                    $resultObj.Error = $_
                    return $resultObj
                }
            }

            $resultObj.MailContactCreated = $true

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
                    $null = $contact.EmailAddresses.Add($addr)

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
            try {
                # Enable as UserMailbox.
                Write-Verbose "Enabling $username as a UserMailbox"

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

            try {
                # Set Audit Logging.
                Write-Verbose "Enabling Audit Logging for $username"
                Set-Mailbox -Identity $username `
                            -AuditEnabled:$true `
                            -AuditLogAgeLimit "90.00:00:00" `
                            -AuditAdmin     Update, Copy, Move, MoveToDeletedItems, SoftDelete, HardDelete, SendAs, SendOnBehalf, MessageBind, Create `
                            -AuditDelegate  Update, Move, MoveToDeletedItems, SoftDelete, HardDelete, SendAs, SendOnBehalf, Create `
                            -AuditOwner     Update, Move, MoveToDeletedItems, SoftDelete, HardDelete `
                            -DomainController $dc `
                            -ErrorAction Stop
            } catch {
                $err =  "An error occurred while enabling audit logging.  The error was: $_"
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
                $key -eq 'FirstName' -or
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
        if ($SendEmailNotification -eq $true -and $EmailNotificationPath -ne $null) {
            if ($MailboxLocation -eq "Local") {
                # Queue up a welcome email to the local mailbox.
                $emailDetails = New-EmailDetailsObject `
                                    -Identity $User.Name `
                                    -Address $User.PrimarySmtpAddress.ToString() `
                                    -Subject "You now have a JMU Exchange E-mail Account"
                $emailDetails.MessageBody = Resolve-TemplatedEmail `
                                    -FilePath $LocalWelcomeEmailTemplate `
                                    -ToBase64String `
                                    -TemplateSubstitutions @{
                                        "[[FirstName]]" = $savedAttributes["FirstName"]
                                        "[[EmailAddress]]" = $User.PrimarySmtpAddress.ToString()
                                    }
                $null = $EmailNotifications.Add($emailDetails)

                if ($contact -ne $null) {
                    # Queue up a notification email to the remote mailbox.
                    $emailDetails = New-EmailDetailsObject `
                                        -Identity $User.Name `
                                        -Address $contact.ExternalEmailAddress.SmtpAddress `
                                        -Subject "You now have a JMU Exchange E-mail Account"
                    $emailDetails.MessageBody = Resolve-TemplatedEmail `
                                        -FilePath $LocalNotificationEmailTemplate `
                                        -ToBase64String `
                                        -TemplateSubstitutions @{
                                            "[[FirstName]]" = $savedAttributes["FirstName"]
                                            "[[EmailAddress]]" = $User.PrimarySmtpAddress.ToString()
                                        }
                    $null = $EmailNotifications.Add($emailDetails)
                }
            } elseif ($MailboxLocation -eq "Remote") {
                if ($contact -eq $null) {
                    # Queue up a welcome email to the remote mailbox.
                    $emailDetails = New-EmailDetailsObject `
                                        -Identity $User.Name `
                                        -Address $User.ExternalEmailAddress.SmtpAddress `
                                        -Subject "You now have a JMU Dukes E-mail Account"
                    $emailDetails.MessageBody = Resolve-TemplatedEmail `
                                        -FilePath $RemoteWelcomeEmailTemplate `
                                        -ToBase64String `
                                        -TemplateSubstitutions @{
                                            "[[FirstName]]" = $savedAttributes["FirstName"]
                                            "[[EmailAddress]]" = $User.ExternalEmailAddress.SmtpAddress
                                        }
                    $null = $EmailNotifications.Add($emailDetails)
                } else {
                    # Queue up a welcome email to the remote mailbox
                    $emailDetails = New-EmailDetailsObject `
                                        -Identity $User.Name `
                                        -Address $contact.ExternalEmailAddress.SmtpAddress `
                                        -Subject "You now have a JMU Dukes E-mail Account"
                    $emailDetails.MessageBody = Resolve-TemplatedEmail `
                                        -FilePath $RemoteWelcomeEmailTemplate `
                                        -ToBase64String `
                                        -TemplateSubstitutions @{
                                            "[[FirstName]]" = $savedAttributes["FirstName"]
                                            "[[EmailAddress]]" = $contact.ExternalEmailAddress.SmtpAddress
                                        }
                    $null = $EmailNotifications.Add($emailDetails)

                    # Queue up a notification email to the local mailbox.
                    $emailDetails = New-EmailDetailsObject `
                                        -Identity $User.Name `
                                        -Address $User.PrimarySmtpAddress.ToString() `
                                        -Subject "You now have a JMU Dukes E-mail Account"
                    $emailDetails.MessageBody = Resolve-TemplatedEmail `
                                        -FilePath $RemoteNotificationEmailTemplate `
                                        -ToBase64String `
                                        -TemplateSubstitutions @{
                                            "[[FirstName]]" = $savedAttributes["FirstName"]
                                            "[[EmailAddress]]" = $contact.ExternalEmailAddress.SmtpAddress
                                        }
                    $null = $EmailNotifications.Add($emailDetails)
                }
            }
        } else {
            Write-Verbose "Not sending emails"
        }

        $resultObj.ProvisioningSuccessful = $true
        $resultObj.Error = "$MailboxLocation mailbox provisioned."
        return $resultObj
    } # end 'PROCESS{}'

    # This section executes only once, after the pipeline.
    END {
        if ($emailNotifications -ne $null -and $emailNotifications.Count -gt 0 -and
                [String]::IsNullOrEmpty($emailNotificationFile) -eq $false) {
            Write-Verbose "Writing out $emailNotificationFile"
            $emailNotifications | Export-Csv -NoTypeInformation -Encoding ASCII -Path $emailNotificationFile
        }
        Write-Verbose "Cleaning up"
    } # end 'END{}'
}
