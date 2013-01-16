################################################################################
# 
# Copyright (c) 2012 Seth Wright <wrightst@jmu.edu>
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

function Deprovision-User {
    [CmdletBinding(SupportsShouldProcess=$true,
        ConfirmImpact="High")]

    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [Alias("Identity")]
            [Alias("Name")]
            # The user to be deprovisioned
            $User,
        
            [Parameter(Mandatory=$false)]
            # The target email address for the user.  Using this parameter is
            # only valid when you want to "deprovision" a UserMailbox into a
            # MailUser.
            [string]
            $ExternalEmailAddress,

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
        
        Write-Verbose "Initialization complete."
    } # end 'BEGIN{}'

# This section executes for each object in the pipeline.
    PROCESS {
        Write-Verbose "Beginning deprovisioning process for `"$User`""
        # Was a username passed to us?  If not, bail.
        if ([String]::IsNullOrEmpty($User)) { 
            Write-Error "User was null."
            return
        }
        $userName = $User

        Write-Verbose "Finding user in Active Directory"
        $objUser = Get-User $User -ErrorAction SilentlyContinue

        if ($objUser -eq $null) {
            Write-Error "$User is not a valid user in Active Directory."
            return
        }
        $userName = $objUser.SamAccountName
        Write-Verbose "sAmAccountName:  $userName"

        # Save these off because Exchange blanks them out...
        $displayName = $objUser.DisplayName
        $displayNamePrintable = $objUser.SimpleDisplayName
        Write-Verbose "displayName:  $displayName"
        Write-Verbose "displayNamePrintable: $displayNamePrintable"

        Write-Verbose "$userName is a $($objUser.RecipientTypeDetails)"
        switch ($objUser.RecipientTypeDetails) {
            'User' { 
                Write-Error "$userName is already a User, and cannot be deprovisioned further."
                return
            }
            'MailUser' { 
                if ([String]::IsNullOrEmpty($ExternalEmailAddress) -eq $false) {
                    Write-Error "The ExternalEmailAddress parameter is only valid to deprovision a UserMailbox to a MailUser"
                    return
                }

                Write-Verbose "Disabling $userName as a MailUser"
                $error.Clear()
                Disable-MailUser -Identity $userName -Confirm:$Confirm `
                    -DomainController $dc -ErrorAction SilentlyContinue

                if (![String]::IsNullOrEmpty($error[0])) {
                    Write-Error $error[0]
                    return
                }
                break
            }
            'UserMailbox' {

                # Save off any extra attributes that we'll need when we
                # make this UserMailbox a MailUser.
                Write-Verbose "Saving Exchange-specific attributes in order to reapply them later"
                $recip = Get-Mailbox $userName
                $legacyExchangeDn = $recip.LegacyExchangeDN
                $emailAddresses = $recip.EmailAddresses

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

                Write-Verbose "Disabling $userName as a UserMailbox"
                $error.Clear()
                Disable-Mailbox -Identity $objUser.DistinguishedName -Confirm:$Confirm `
                    -DomainController $dc -ErrorAction SilentlyContinue

                if (![String]::IsNullOrEmpty($error[0])) {
                    Write-Error $error[0]
                    return
                }

                foreach ($i in 1..12) {
                    Write-Verbose "Waiting for Active Directory replication to finish..."
                    Start-Sleep 5

                    $objUser = Get-User $userName `
                        -ErrorAction SilentlyContinue `
                        -DomainController $dc

                    if ($objUser.RecipientTypeDetails -eq "User") {
                        Write-Verbose "Replication has finished."
                        break
                    }
                }

                if ($objUser -eq "UserMailbox") {
                    Write-Error "An error occurred disabling $userName as a UserMailbox."
                    return
                }

                Write-Verbose "$userName has been disabled as a UserMailbox"
                break
                }
            default {
                Write-Error "$userName is a $($objUser.RecipientTypeDetails) object, and I have no idea how to deprovision it."
                return
            }
        }

        if ($objUser.RecipientTypeDetails -eq 'User' -and 
            [String]::IsNullOrEmpty($ExternalEmailAddress) -eq $false) {
            Write-Verbose "Enabling user `"$userName`" as a MailUser with forwarding address of $ExternalEmailAddress"
            $error.Clear()
            Write-Verbose "Calling Provision-User"
            Provision-User -User $userName -MailUser `
                -ExternalEmailAddress $ExternalEmailAddress `
                -EmailAddresses $EmailAddresses `
                -Verbose:$Verbose -DomainController $dc

            if (![String]::IsNullOrEmpty($error[0])) {
                Write-Error $error[0]
                return
            }

            Write-Verbose "Provision-User finished"

            # If the object was a UserMailbox before, reset all of the Custom
            # Attributes.
            if ($objUser.RecipientTypeDetails -eq "UserMailbox") {
                Write-Verbose "Resetting CustomAttributes (extensionAttribute*) 1 - 15 to saved values"
                Set-MailUser -Identity $userName `
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
        }

        Write-Verbose "Resetting displayName and displayNamePrintable attributes to saved values"
        # Set various attributes now that Exchange has helpfully removed them.
        $error.Clear()
        Set-User -Identity $userName -DisplayName $displayName `
            -SimpleDisplayName $displayNamePrintable `
            -DomainController $dc `
            -ErrorAction SilentlyContinue

        if (![String]::IsNullOrEmpty($error[0])) {
            Write-Error $error[0]
            return
        }


        Write-Verbose "Ending deprovisioning process for `"$User`""
    } # end 'PROCESS{}'

# This section executes only once, after the pipeline.
    END {
        Write-Verbose "Cleaning up"
    } # end 'END{}'
}
