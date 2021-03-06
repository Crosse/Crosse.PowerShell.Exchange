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

function New-Resource {
    [CmdletBinding(SupportsShouldProcess=$true,
            ConfirmImpact="High")]

    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The name of the new resource as it will appear in the GAL.
            $DisplayName,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [Alias("Owner")]
            [string]
            # The owner of the new resource.
            $Delegate,

            [switch]
            # This resource is a room resource.
            $Room,

            [switch]
            # This resource is an equipment resource.
            $Equipment,

            [switch]
            # This resource is a shared mailbox or a shared calendar.
            $Shared,

            [switch]
            # This resource is a shared calendar.
            $Calendar,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # For shared mailboxes, the SMTP address to set.
            $PrimarySmtpAddress = "",

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The base "ExchangeObjects" OU.  Defaults to "<current.domain.local>/ExchangeObjects"
            $BaseDN,

            [switch]
            # Whether to email the owner.  Default is true.
            $EmailOwner=$true,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The SMTP server used to when sending email.
            $SmtpServer,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The From address used when sending email.
            $From,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string[]]
            # An array of email addresses to BCC when sending email to owners.
            $Bcc,

            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The domain controller to target for all operations.
            $DomainController
          )

    BEGIN {
        Write-Verbose "Performing initialization actions."

        $Confirm = ($PSBoundParameters["Confirm"] -eq $null -or $PSBoundParameters["Confirm"].ToBool())
        $Verbose = ($PSBoundParameters["Verbose"] -eq $null -or $PSBoundParameters["Verbose"].ToBool())

        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        $controllers = @($domain.DomainControllers | Foreach-Object { $_.Name.ToLower() })
        $controllersCount = $controllers.Count
        Write-Verbose "Found $controllersCount domain controllers."

        if ([String]::IsNullOrEmpty($DomainController)) {
            $ForceRediscovery = [System.DirectoryServices.ActiveDirectory.LocatorOptions]::ForceRediscovery
            while ($dc -eq $null) {
                Write-Verbose "Finding a global catalog to use for this operation"
                $controller = $domain.FindDomainController($ForceRediscovery)
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
    } # end 'BEGIN{}'

    PROCESS {
        if ( !($Room -or $Equipment -or $Shared) ) {
            Write-Error "Please specify either -Room, -Equipment, -Shared, or -Calendar"
            return
        }

        if (($Room -and $Equipment) -or ($Room -and $Shared) -or ($Equipment -and $Shared)) {
            Write-Error "Please specify only one of -Room, -Equipment, or -Shared"
            return
        }

        if ( $Shared -and !$Calendar -and ($PrimarySmtpAddress -eq "") ) {
            Write-Error "Please specify the PrimarySmtpAddress"
            return
        }

        $Owner = Get-Mailbox $Delegate
        if ($Owner -eq $null) {
            Write-Error "Could not find owner $Delegate"
            return
        }

        if ($Shared -and ![String]::IsNullOrEmpty($PrimarySMTPAddress)) {
            $tempResource = Get-Recipient -Anr $PrimarySMTPAddress -DomainController $dc
            if ($tempResource -ne $null) {
                Write-Error "The PrimarySmtpAddress already exists on object `"$($tempResource.Name)`""
                return
            }
        }
        if ([String]::IsNullOrEmpty($BaseDN)) {
            $ou = $domain.Name + "/ExchangeObjects"
        } else {
            $ou = $BaseDN
        }

        if  ( $Room ) {
            $ou += "/Resources/Rooms"
        } elseif ( $Equipment ) {
            $ou += "/Resources/Equipment"
        } elseif ( $Shared -and !$Calendar) {
            $ou += "/SharedMailboxes"
        } elseif ( $Shared -and $Calendar) {
            $ou += "/Resources/SharedCalendars"
        }

        $Name  = $DisplayName
        $alias = $DisplayName
        $alias = $alias.Replace('Conference Room', 'ConfRoom')
        $alias = $alias.Replace('Lecture Hall', 'LectureHall')
        $alias = $alias.Replace(' Hall', '')
        $alias = $alias -replace '[\s-#&()]+', '_'

        if ($Shared -and !$Calendar) {
            $alias += "_Mailbox"
        }

        $newMailboxParams = @{
            DomainController = $dc
            Name = $Name
            Alias = $Alias
            Displayname = $DisplayName
            UserPrincipalName = "$($alias)@ad.jmu.edu"
            OrganizationalUnit = $ou
            ManagedFolderMailboxPolicy = "Default Managed Folder Policy"
            ManagedFolderMailboxPolicyAllowed = $true
        }

        if ( $Room ) {
            $newMailboxParams["Room"] = $true
        } elseif ( $Equipment ) {
            $newMailboxParams["Equipment"] = $true
        } elseif ( $Shared ) {
            $newMailboxParams["Shared"] = $true
        }

        $resource = New-Mailbox -ErrorAction Stop @newMailboxParams

        if ( !$resource) {
            Write-Error "Could not find $alias in Active Directory."
            return
        }

        if ($OtherMailboxGuid -ne $null) {
            Write-Verbose "Setting otherMailbox to $OtherMailboxGuid"
            $null = Set-ADAttribute -DomainController $dc `
                        -Identity $resource.UserPrincipalName `
                        -Attribute otherMailbox `
                        -Value $OtherMailboxGuid.ToString('b')
        }

        $timeout = 60
        $waitingOn = $controllers.Count
        $foundOn = "no domain controllers yet"
        $activity =  "Waiting for AD replication..."
        $status = "Waiting on $waitingOn domain controllers"
        for ($i = 0; $i -lt $timeout; $i++) {
            [Int32]$pct = [Math]::Round($i*100/$timeout, 0)
            Write-Progress -Activity $activity -Status $status `
                -CurrentOperation "Replicated to $foundOn" `
                -SecondsRemaining ($timeout - $i) -PercentComplete $pct

            if ($waitingOn -gt 0) {
                $waitingOn = $controllers.Count
                $replicated = @()
                foreach ($c in $controllers) {
                    $mbx = Get-ADAttribute -DomainController $c `
                                           -Identity $resource.UserPrincipalName `
                                           -Attribute msExchMailboxGuid `
                                           -ErrorAction SilentlyContinue
                    if ($mbx -ne $null -and $mbx.msExchMailboxGuid -ne $null) {
                        $replicated += $c.ToLower() -replace ".$($domain.Name.ToLower())", ""
                        $waitingOn--
                    }
                }
                $foundOn = $replicated -join ", "
            }
            Start-Sleep 1
        }
        Write-Progress -Activity "Waiting for AD replication..." -Status "Completed" -PercentComplete 100 -Completed

        Write-Verbose "Sending a message to the resource to initialize the mailbox."
        Send-MailMessage -From $From -To "$($resource.PrimarySMTPAddress)" -Subject "disregard" -Body "disregard" -SmtpServer $SmtpServer

        if ($Equipment -or $Room) {
            # If this is a Resource mailbox and not a Shared mailbox...

            # Set the default mailbox quota limits.
            $timeout = 60 ; $i = 0
            $activity = "Setting mailbox defaults"
            $status = "Setting mailbox quota defaults"
            for ($i = 0; $i -lt $timeout; $i++) {
                [Int32]$pct = [Math]::Round($i*100/$timeout, 0)
                Write-Progress -Activity $activity -Status $status `
                    -CurrentOperation "Attempt $i" `
                    -SecondsRemaining ($timeout - $i) -PercentComplete $pct

                try {
                    Write-Verbose "Attempting to set mailbox quota and auditing defaults"

                    $auditOwner = @("Update", "Move", "MoveToDeletedItems", "SoftDelete", "HardDelete")
                    $auditDelegate = $auditOwner + @("SendAs", "SendOnBehalf", "Create")
                    $auditAdmin = $auditDelegate + @("Copy", "MessageBind")
                    $resource | Set-Mailbox -DomainController $dc `
                                -IssueWarningQuota 90MB `
                                -ProhibitSendQuota 95MB `
                                -ProhibitSendReceiveQuota 100MB `
                                -UseDatabaseQuotaDefaults:$false `
                                -AuditEnabled:$true `
                                -AuditLogAgeLimit "90.00:00:00" `
                                -AuditOwner $auditOwner `
                                -AuditDelegate $auditDelegate `
                                -AuditAdmin $auditAdmin `
                                -ErrorAction Stop
                    break
                } catch {
                    #Write-Warning $_
                    Start-Sleep 1
                }
            }
            Write-Progress -Activity $activity -Status $status -Completed
            if ($i -eq $timeout) {
                Write-Error "Unable to set quota defaults!"
            }

            $timeout = 120 ; $i = 0
            $activity = "Setting mailbox defaults"
            $status = "Setting calendaring defaults"
            # Set the default calendar settings on the resource.
            for ($i = 0; $i -lt $timeout; $i++) {
                [Int32]$pct = [Math]::Round($i*100/$timeout, 0)
                Write-Progress -Activity $activity -Status $status `
                    -CurrentOperation "Attempt $i" `
                    -SecondsRemaining ($timeout - $i) -PercentComplete $pct

                try {
                    Write-Verbose "Attempting to set default calendar settings"
                    $resource | Set-CalendarProcessing -DomainController $dc `
                                -AllRequestOutOfPolicy:$True -AutomateProcessing AutoAccept `
                                -BookingWindowInDays 365 -ResourceDelegates $Owner `
                                -ErrorAction Stop
                    Write-Verbose "Took $i seconds for the resource to become available."
                    break
                } catch {
                    #Write-Warning $_
                    Start-Sleep 1
                }
            }
            Write-Progress -Activity $activity -Status $status -Completed
            if ($i -eq $timeout) {
                Write-Error "Unable to set calendar defaults!"
            }
        } elseif ( $Shared -and !$Calendar ) {
                # Set the target mailbox's EmailAddresses property to include the PrimarySMTPAddress
                # specified on the command line.
                $emailAddresses = $resource.EmailAddresses
                # Add the @jmu.edu address as the PrimarySmtpAddress
                if (!($emailAddresses.Contains("SMTP:$($PrimarySmtpAddress)")) ) {
                    $emailAddresses.Add("SMTP:$($PrimarySmtpAddress)")
                }

                $proxyAddress = $PrimarySmtpAddress.Replace("@jmu.edu", "@ad.jmu.edu")
                if (!($emailAddresses.Contains("smtp:$($proxyAddress)")) ) {
                    $emailAddresses.Add("smtp:$($proxyAddress)")
                }
                $resource | Set-Mailbox -EmailAddressPolicyEnabled:$False -EmailAddresses $emailAddresses `
                            -DomainController $dc
        }

        Write-Verbose "Adding $Owner as a delegate"
        Add-ResourceDelegate -DomainController $dc `
            -ResourceIdentity $resource -Delegate $Owner `
            -EmailDelegate:$EmailOwner `
            -SmtpServer $SmtpServer -From $From -Bcc $Bcc `
            -Confirm:$Confirm -Verbose:$Verbose

        return (Get-Mailbox -DomainController $dc -Identity $resource)
    }
}
