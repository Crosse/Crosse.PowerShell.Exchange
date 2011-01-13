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

        Write-Verbose "Initialization complete."
    } # end 'BEGIN{}'

    
# This section executes for each object in the pipeline.
    PROCESS {
        Write-Verbose "Beginning provisioning process for `"$Identity`""

        $User = $null
        if ($Identity -is [System.String]) {
            $User = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction SilentlyContinue
            if ($User -eq $null) {
                Write-Error "$Identity is not a valid user in Active Directory."
                return
            }
            Write-Verbose "Found user $User in Active Directory"
        } elseif ($Identity -is [Microsoft.Exchange.Data.Directory.Management.User] -or
                  $Identity -is [Microsoft.Exchange.Data.Directory.Management.MailEnabledOrgPerson] -or
                  $Identity -is [Microsoft.Exchange.Data.Directory.ADObjectId] -or
                  $Identity -is [Microsoft.Exchange.Data.Directory.ADObject]) {
            $User = Get-User -Identity $Identity.Name -DomainController $DomainController
        } else {
            Write-Error "$Identity is of non-valid type $($Identity.GetType()) for this operation."
            return
        }

        Write-Verbose "RecipientTypeDetails:  $($User.RecipientTypeDetails)"
        $username = $User.SamAccountName

        # We don't process disabled users, at least not right now.
        if ($User.RecipientTypeDetails -eq 'DisabledUser') {
            Write-Error "$username is disabled in Active Directory."
            return
        }

        # There is no sense in processing a UserMailbox if "Local" 
        # was specified, since it already has a mailbox.
        if ($User.RecipientTypeDetails -eq 'UserMailbox' -and $MailboxLocation -eq 'Local') {
            Write-Error "$username already has a local mailbox."
            return
        }

        # Another condition that doesn't make sense:  a MailUser getting a 
        # "Remote" mailbox.
        if ($User.RecipientTypeDetails -eq 'MailUser' -and $MailboxLocation -eq 'Remote') {
            Write-Error "$username already has a remote mailbox."
            return
        }

        $desc = "Provision $MailboxLocation Mailbox for `"$username`""
        $caption = $desc
        $warning = "Are you sure you want to perform this action?`n"
        $warning += "This will give the user `"$username`" a "
        $warning += "$($MailboxLocation.ToLower()) mailbox."

        if (!$PSCmdlet.ShouldProcess($desc, $warning, $caption)) {
            return
        }
            
        # Attributes that all user objects have.
        $savedAttributes = New-Object System.Collections.Hashtable
        $savedAttributes["DisplayName"] = $User.DisplayName
        $savedAttributes["SimpleDisplayName"] = $User.SimpleDisplayName

        # Save off attributes that tend to get blanked out.
        if ($User.RecipientTypeDetails -eq 'MailUser' -or
            $User.RecipientTypeDetails -eq 'UserMailbox') {

            $recipient = $null
            if ($User.RecipientTypeDetails -eq 'MailUser') {
                $recipient = Get-MailUser $username -DomainController $DomainController
                if ($recipient -eq $null) {
                    Write-Error "Could not perform Get-MailUser on $username"
                    return
                }
            } elseif ($User.RecipientTypeDetails -eq 'UserMailbox') { 
                $recipient = Get-Mailbox $username -DomainController $DomainController
                if ($recipient -eq $null) {
                    Write-Error "Could not perform Get-Mailbox on $username"
                    return
                }
            }

            # Attributes that only MailUsers and UserMailboxes have.
            $savedAttributes["ExternalEmailAddress"] = $recipient.ExternalEmailAddress
            $savedAttributes["LegacyExchangeDN"] = $recipient.LegacyExchangeDN
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
            $savedAttributes["CustomAttribute15"] = $recipient.CustomAttribute15
        }

        foreach ($key in ($savedAttributes.Keys | Sort-Object)) {
            Write-Verbose "$($key):`t$($savedAttributes[$key])"
        }

        # MailUser --> UserMailbox:  user must first be disabled as a 
        # MailUser, then it can be enabled as a mailbox.
        if ($User.RecipientTypeDetails -eq 'MailUser') {
            Write-Verbose "Disabling $username as a MailUser"
            $error.Clear()
            Disable-MailUser -Identity $username -DomainController $DomainController -Confirm:$false

            # Object should now just be a "User".  Verify.
            $User = Get-User -Identity $username -DomainController $DomainController -ErrorAction SilentlyContinue
            if ($User.RecipientTypeDetails -ne 'User') {
                Write-Error "An error occurred while running Disable-MailUser.  The error was:  $error[0]"
                Write-Host "Attribute Dump:"
                foreach ($key in ($savedAttributes.Keys | Sort-Object)) {
                    Write-Host "$($key):`t$($savedAttributes[$key])"
                }
                return
            }
        }

        # Enable the mailbox.
        Write-Verbose "Enabling `"$username`" as a UserMailbox"
        Enable-Mailbox  -Identity $username `
                        -ManagedFolderMailboxPolicy "Default Managed Folder Policy" `
                        -ManagedFolderMailboxPolicyAllowed:$true `
                        -DomainController $DomainController `
                        -ErrorAction SilentlyContinue | Out-Null

        $User = Get-User -Identity $username -DomainController $DomainController -ErrorAction SilentlyContinue
        if ($User.RecipientTypeDetails -ne 'UserMailbox') {
            Write-Error "An error occurred while running Enable-Mailbox.  The error was: $error[0]"
            Write-Host "Attribute Dump:"
            foreach ($key in ($savedAttributes.Keys | Sort-Object)) {
                Write-Host "$($key):`t$($savedAttributes[$key])"
            }
            return
        }

        # Reapply any saved attributes.
        $cmd = "Set-Mailbox -Identity $Identity -DomainController $DomainController "
        foreach ($key in $savedAttributes.Keys) {
            if ($key -eq 'LegacyExchangeDN' -or $key -eq 'ExternalEmailAddress') {
                continue
            }
            $cmd += "-$($key) `"$($savedAttributes[$key])`" "
        }
        Invoke-Expression $cmd
    }
}
