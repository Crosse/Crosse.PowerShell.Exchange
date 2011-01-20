################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  Provisions resources in Exchange for JMU
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

            [Parameter(Mandatory=$false)]
            # Whether to force-create a mailbox for a user, even if they would
            # not normally be a candidate for a mailbox
            [switch]
            $Force,

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
            [string]
            # The domain controller to use for all operations.
            $DomainController,

            [Parameter(Mandatory=$false)]
            [string]
            # Where to create MailContact objects, if necessary.
            $MailContactOrganizationalUnit="ad.test.jmu.edu/ExchangeObjects/MailContacts"
        )

# This section executes only once, before the pipeline.
    BEGIN {
        Write-Verbose "Performing initialization actions."

        if ([String]::IsNullOrEmpty($DomainController)) {
            $DomainController = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController().Name

            if ($DomainController -eq $null) {
                Write-Error "Could not find a domain controller to use for the operation."
                return
            }
        }
        Write-Verbose "Using Domain Controller $DomainController"
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
            ProvisioningSuccessful  = $false
            Error                   = $null
        }
            
        $User = $null
        try {
            $User = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop
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
            $err = "$username is disabled in Active Directory.  Not performing any operations."
            Write-Error $err
            $resultObj.Error = $err
            return $resultObj
        }

        # There is no sense in processing a UserMailbox if "Local" 
        # was specified, since it already has a mailbox.
        if ($User.RecipientTypeDetails -eq 'UserMailbox' -and $MailboxLocation -eq 'Local') {
            $err = "$username already has a local mailbox."
            Write-Error $err
            $resultObj.Error = $err
            return $resultObj
        }

        # Another condition that doesn't make sense:  a MailUser getting a 
        # "Remote" mailbox.
        if ($User.RecipientTypeDetails -eq 'MailUser' -and $MailboxLocation -eq 'Remote') {
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

        if ([String]::IsNullOrEmpty($ExternalEmailAddress) -eq $false) {
            $savedAttributes["ExternalEmailAddress"] = $ExternalEmailAddress
        }
        
        # Save some attributes that tend to get blanked out.
        if ($User.RecipientTypeDetails -eq 'MailUser' -or
            $User.RecipientTypeDetails -eq 'UserMailbox') {

            $recipient = $null
            if ($User.RecipientTypeDetails -eq 'MailUser') {
                try {
                    $recipient = Get-MailUser $username -DomainController $DomainController -ErrorAction Stop
                } catch {
                    $err =  "Could not perform Get-MailUser on $username:  $_"
                    Write-Error $err
                    $resultObj.Error = $err
                    return $resultObj
                }

                # Attributes that only MailUsers have.
                $savedAttributes["ExternalEmailAddress"] = $recipient.ExternalEmailAddress
                $savedAttributes["LegacyExchangeDN"] = $recipient.LegacyExchangeDN
            } elseif ($User.RecipientTypeDetails -eq 'UserMailbox') {
                try {
                    $recipient = Get-Mailbox $username -DomainController $DomainController -ErrorAction Stop
                } catch {
                    $err = "Could not perform Get-Mailbox on $username:  $_"
                    Write-Error $err
                    $resultObj.Error = $err
                    return $resultObj
                }
            }

            # Attributes that both MailUsers and UserMailboxes have.
            $savedAttributes["CustomAttribute1"] = $recipient.CustomAttribute1
            $savedAttributes["CustomAttribute2"] = $recipient.CustomAttribute2
            $savedAttributes["CustomAttribute3"] = $recipient.CustomAttribute3
            $savedAttributes["CustomAttribute4"] = $recipient.CustomAttribute4
            $savedAttributes["CustomAttribute5"] = $recipient.CustomAttribute5
            $savedAttributes["CustomAttribute6"] = $recipient.CustomAttribute6
            $savedAttributes["CustomAttribute7"] = $recipient.CustomAttribute7
            $savedAttributes["CustomAttribute8"] = $recipient.CustomAttribute8
            $savedAttributes["CustomAttribute9"] = $recipient.CustomAttribute9
            $savedAttributes["CustomAttribute10"] = $recipient.CustomAttribute10
            $savedAttributes["CustomAttribute11"] = $recipient.CustomAttribute11
            $savedAttributes["CustomAttribute12"] = $recipient.CustomAttribute12
            $savedAttributes["CustomAttribute13"] = $recipient.CustomAttribute13
            $savedAttributes["CustomAttribute14"] = $recipient.CustomAttribute14
        }

            # This will be the "last touched time" attribute and should
            # be set for object.
            $savedAttributes["CustomAttribute15"] = "LastProvisioned: $(Get-Date -Format u)"

        foreach ($key in ($savedAttributes.Keys)) {
            Write-Verbose "$($key):`t$($savedAttributes[$key])"
        }

        if ($User.RecipientTypeDetails -eq 'MailUser') {
            # MailUser --> UserMailbox:  user must first be disabled as a 
            # MailUser, then it can be enabled as a mailbox.
            Write-Verbose "Disabling $username as a MailUser"

            try {
                Disable-MailUser `
                    -Identity $username `
                    -DomainController $DomainController `
                    -Confirm:$false `
                    -ErrorAction Stop
            } catch {
                $err = "Could not run Disable-MailUser for user $username ($_)."
                Write-Error $err
                $resultObj.Error = $err
                return $resultObj
            }

            # Object should now just be a "User".
            $User = Get-User -Identity $username -DomainController $DomainController -ErrorAction Stop
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
            try {
                $contact = New-MailContact `
                                -Name "$($username)-mc" `
                                -Alias "$($username)-mc" `
                                -DisplayName "$($User.DisplayName) (Dukes)" `
                                -FirstName "$($User.FirstName)" `
                                -LastName "$($User.LastName)" `
                                -ExternalEmailAddress $savedAttributes["ExternalEmailAddress"] `
                                -OrganizationalUnit $MailContactOrganizationalUnit `
                                -DomainController $DomainController `
                                -ErrorAction Stop
            } catch {
                $err =  "Could not create contact for $username.  The error was:  $_"
                Write-Error $err
                $resultObj.Error = $_
                return $resultObj
            }

            $resultObj.MailContactCreated = $true

            if ([String]::IsNullOrEmpty($savedAttributes["LegacyExchangeDN"]) -eq $false) {
                Write-Verbose "Adding LegacyExchangeDN to new MailContact as an X.400 address"
                $addr = "X.400:" + $savedAttributes["LegacyExchangeDN"]
                $contact.EmailAddresses.Add($addr)

                try {
                    $contact = Set-MailContact `
                                    -Identity $contact.Identity `
                                    -EmailAddresses $contact.EmailAddresses `
                                    -DomainController $DomainController `
                                    -ErrorAction Stop
                } catch {
                    $w = "An error occurred while adding " + $savedAttributes["LegacyExchangeDN"]
                    $w += " as an X.400 address to the MailContact for $username.  "
                    $w += "You will need to add this manually.  The error was:  $_"
                    Write-Warning $w
                    $resultObj.Error = $w
                }
                               
                $contact = Get-MailContact -Identity $contact.Identity -DomainController $DomainController
            }
        }

        # At this point, the following is true:
        # * If "Remote" was specified and the object is a UserMailbox, a 
        #   MailContact has been created and that's all that needs to happen.
        #   
        # * If "Local" was specified--no matter what type the object was as 
        #   the start--then the object will be enabled as a UserMailbox.
        #   
        # * If "Remote" was specified and the object is a User, the object will
        #   be enabled as a MailUser.

        if ($MailboxLocation -eq 'Local') {
            # Enable as UserMailbox.
            Write-Verbose "Enabling $username as a UserMailbox"

            try {
                $User = Enable-Mailbox `
                            -Identity $username `
                            -ManagedFolderMailboxPolicy "Default Managed Folder Policy" `
                            -ManagedFolderMailboxPolicyAllowed:$true `
                            -DomainController $DomainController `
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
                            -DomainController $DomainController `
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
        if ($User.RecipientTypeDetails -eq 'MailUser' -or $User.RecipientTypeDetails -eq 'UserMailbox') {
            if ($User.RecipientTypeDetails -eq 'MailUser') {
                $cmd += "Set-MailUser "
            } elseif ($User.RecipientTypeDetails -eq 'UserMailbox') {
                $cmd += "Set-Mailbox "
            }
            $cmd += "-Identity $Identity -DomainController $DomainController "

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
                Invoke-Expression $cmd
            } catch {
                $err =  "An error occurred while reapplying saved attributes.  The error was: $_"
                Write-Error $err
                $resultObj.Error = $err
                return $resultObj
            }
        }
        $resultObj.ProvisioningSuccessful = $true
        return $resultObj
    } # end 'PROCESS{}'

# This section executes only once, after the pipeline.
    END {
        Write-Verbose "Cleaning up"
    } # end 'END{}'
}
