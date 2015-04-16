################################################################################
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
################################################################################

function New-Resource {
    [CmdletBinding(SupportsShouldProcess=$true,
            ConfirmImpact="High")]

    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $DisplayName, 

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Owner, 
            
            [Parameter(ParameterSetName='Resource',
                Mandatory=$true)]
            [switch]
            $Room,
            
            [switch]
            $Equipment,
            
            [switch]
            $Shared,
            
            [switch]
            $Calendar,
            
            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            $PrimarySmtpAddress,
            
            [Parameter(Mandatory=$false)]
            [ValidateNotNullOrEmpty()]
            [string]
            $BaseDN = "ad.jmu.edu/ExchangeObjects",

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
            $Bcc = @("wrightst@jmu.edu", "richa3jb@jmu.edu", "eckardsl@jmu.edu"),

            [switch]
            $EmailOwner = $true,

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
    } # end 'BEGIN{}'

    # This section executes for each object in the pipeline.
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

        try {
            $objUser = Get-User $Owner -DomainController $dc -ErrorAction Stop
        } catch {
            Write-Error -ErrorRecord $_
            return
        }

        if ($Shared) {
            $tempResource = Get-Recipient -Anr $PrimarySMTPAddress -DomainController $dc
            if ($tempResource -ne $null) {
                Write-Error "The PrimarySmtpAddress already exists:"
                $tempResource
                return
            }
        }

        $ou = $BaseDN
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
        $alias = $alias.Replace(' - ', '_')
        $alias = $alias.Replace(' ', '_')
        $alias = $alias.Replace('#', '')
        if ($Shared -and !$Calendar) {
            $alias += "_Mailbox"
        }

        $cmd  = "New-Mailbox -DomainController $dc "
        $cmd += "-OrganizationalUnit `"$ou`" -Name `"$Name`" -Alias `"$alias`" -UserPrincipalName "
        $cmd += "`"$($alias)@ad.jmu.edu`" -DisplayName `"$DisplayName`""
        $cmd += "-ManagedFolderMailboxPolicy 'Default Managed Folder Policy' -ManagedFolderMailboxPolicyAllowed:`$true "

        $error.Clear()

        if ( $Room ) {
            $cmd += " -Room"
        } elseif ( $Equipment ) {
            $cmd += " -Equipment"
        } elseif ( $Shared ) {
            $cmd += " -Shared"
        }

        Invoke-Expression($cmd)

        if (!([String]::IsNullOrEmpty($error[0]))) {
            return
        }

        $resource = Get-Mailbox -DomainController $dc -Identity "$DisplayName"

        if ( !$resource) {
            Write-Error "Could not find $alias in Active Directory."
            return
        }

        # Send a message to the mailbox.  Somehow this helps...but sleep first.
        Write-Host "`nBlocking for 60 seconds for the mailbox to be created:"
        foreach ($i in 1..60) { 
            if ($i % 10 -eq 0) {
                Write-Host -NoNewLine "!"
            } else {
                Write-Host -NoNewLine "."
            }
            Start-Sleep 1
        }

        Write-Host "`ndone.`nSending a message to the resource to initialize the mailbox."
        Send-MailMessage -From $From -To "$($resource.PrimarySMTPAddress)" -Subject "disregard" -Body "disregard" -SmtpServer $SmtpServer

        if ($Equipment -or $Room) {
            # If this is a Resource mailbox and not a Shared mailbox...
            # Set the default calendar settings on the resource.
            # Unfortunately, this fails if the mailbox isn't fully created yet, so introduce a wait.
            Write-Host "Setting Calendar Settings: "

            foreach ($i in 1..10) {
                $error.Clear()
                $resource | Set-CalendarProcessing -DomainController $dc `
                            -AllRequestOutOfPolicy:$True -AutomateProcessing AutoAccept `
                            -BookingWindowInDays 365 -ErrorAction SilentlyContinue
                if (![String]::IsNullOrEmpty($error[0])) {
                    Write-Host -NoNewLine "."
                    Start-Sleep $i
                } else {
                    Write-Host "done."
                    break
                }
            }
        } elseif ( $Shared -and !$Calendar ) {
                # Set the target mailbox's EmailAddresses property to include the PrimarySMTPAddress
                # specified on the command line.
                $emailAddresses = $resource.EmailAddresses
                # Add the @jmu.edu address as the PrimarySmtpAddress
                if (!($emailAddresses.Contains("SMTP:$($PrimarySmtpAddress)")) ) {
                    $emailAddresses.Add("SMTP:$($PrimarySmtpAddress)")
                }

                $resource | Set-Mailbox -EmailAddressPolicyEnabled:$False -EmailAddresses $emailAddresses `
                            -DomainController $DomainController
        }

        Add-ResourceDelegate -ResourceIdentity $resource -Delegate $objUser -DomainController $dc
    }
}
