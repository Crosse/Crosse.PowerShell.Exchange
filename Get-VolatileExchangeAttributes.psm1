################################################################################
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

function Get-VolatileExchangeAttributes {
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [Alias("Name")]
            # Specifies the user to be provisioned.
            $Identity,

            [Parameter(Mandatory=$false)]
            [string]
            # The domain controller to use for all operations.
            $DomainController
          )

    # This section executes only once, before the pipeline.
    BEGIN {
        Write-Verbose "Performing initialization actions."

        if ([String]::IsNullOrEmpty($DomainController)) {
            $dc = [System.DirectoryServices.ActiveDirectory.Domain]::`
                    GetCurrentDomain().FindDomainController().Name

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
        $resultObj = New-Object PSObject -Property @{
            Identity                = $Identity
            Attributes              = $null
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

        $resultObj.Attributes = $savedAttributes

        return $resultObj
    }
}
