################################################################################
# 
# $Id$
# 
# DESCRIPTION:  Creates a new resource in accordance with JMU's current naming
#               policies, etc.
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

param ( [string]$DisplayName, 
        [string]$Owner, 
        [switch]$Room,
        [switch]$Equipment,
        [switch]$Shared,
        [switch]$Calendar,
        [string]$PrimarySmtpAddress = "",
        [switch]$EmailOwner=$true,
        [switch]$Legacy=$false)

# Change these to suit your environment
$SmtpServer         = "it-exhub.ad.jmu.edu"
$From               = "it-exmaint@jmu.edu"
$Bcc                = "wrightst@jmu.edu, millerca@jmu.edu, najdziav@jmu.edu, eckardsl@jmu.edu"
$Fqdn               = "exchange.jmu.edu"
$BaseDN             = "ad.jmu.edu/ExchangeObjects"

##################################
$cwd                = [System.IO.Path]::GetDirectoryName(($MyInvocation.MyCommand).Definition)
$DomainController   = (gc Env:\LOGONSERVER).Replace('\', '')
if ($DomainController -eq $null) { 
    Write-Warning "Could not determine the local computer's logon server!"
    return
}

if ( $DisplayName -eq '' -or $Owner -eq '') {
    Write-Error "-DisplayName and -Owner are required"
    return
}

if ( !($Room -or $Equipment -or $Shared) ) {
    Write-Error "Please specify either -Room, -Equipment, -Shared, or -Calendar"
    return
}

if (($Room -and $Equipment) -or ($Room -and $Shared) -or ($Equipment -and $Shared)) {
    Write-Error "Please specify only one of -Room, -Equipment, or -Shared"
    return
}

if ( $Shared -and !($Calendar) -and ($PrimarySmtpAddress -eq "") ) {
    Write-Error "Please specify the PrimarySmtpAddress"
    return
}

$Owner = Get-Mailbox $Owner
if ($Owner -eq $null) {
    Write-Error "Could not find owner"
    return
}

if ($Shared) {
    $tempResource = Get-Mailbox -Anr $PrimarySMTPAddress
    if ($tempResource -ne $null) {
        Write-Error "The PrimarySmtpAddress already exists"
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
$alias = $alias.Replace(' ', '_')
if ($Shared -and !$Calendar) {
    $alias += "_Mailbox"
}

if ($Legacy) {
    $Database = & "$cwd\Get-BestDatabase.ps1" IT-ExMbx1
    if ($Database -eq $null) {
        Write-Error "Could not find a suitable database!"
        return
    }
}

$cmd  = "New-Mailbox -DomainController $DomainController "
$cmd += "-OrganizationalUnit `"$ou`" -Name `"$Name`" -Alias `"$alias`" -UserPrincipalName "
$cmd += "`"$($alias)@ad.jmu.edu`" -DisplayName `"$DisplayName`""
$cmd += "-ManagedFolderMailboxPolicy 'Default Managed Folder Policy' -ManagedFolderMailboxPolicyAllowed:`$true "
if ($Legacy) {
    $cmd += "-Database `"$Database`""
}

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

$resource = Get-Mailbox -DomainController $DomainController -Identity "$DisplayName"

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
& "$cwd\Send-Email.ps1" -From $From -To "$($resource.PrimarySMTPAddress)" -Subject "disregard" -Body "disregard" -SmtpServer $SmtpServer

# Grant SendOnBehalfOf rights to the owner:
$resource | Set-Mailbox -DomainController $DomainController `
            -GrantSendOnBehalfTo $owner
# Grant Send-As rights to the owner:
$resource | Add-ADPermission -ExtendedRights "Send-As" -User $owner `
            -DomainController $DomainController
# Give the owner Full Access to the resource:
$resource | Add-MailboxPermission -DomainController $DomainController `
            -AccessRights FullAccess -User $Owner

if ($Equipment -or $Room) {
    # If this is a Resource mailbox and not a Shared mailbox...
    # Set the default calendar settings on the resource.
    # Unfortunately, this fails if the mailbox isn't fully created yet, so introduce a wait.
    Write-Host "Setting Calendar Settings: "

    foreach ($i in 1..10) {
        $error.Clear()
        $resource | Set-CalendarProcessing -DomainController $DomainController `
                    -AllRequestOutOfPolicy:$True -AutomateProcessing AutoAccept `
                    -BookingWindowInDays 365 -ResourceDelegates $owner `
                    -ErrorAction SilentlyContinue
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

        $proxyAddress = $PrimarySmtpAddress.Replace("@jmu.edu", "@ad.jmu.edu")
        if (!($emailAddresses.Contains("smtp:$($proxyAddress)")) ) {
            $emailAddresses.Add("smtp:$($proxyAddress)")
        }
        $resource | Set-Mailbox -EmailAddressPolicyEnabled:$False -EmailAddresses $emailAddresses `
                    -DomainController $DomainController
}

if ($EmailOwner) {
    $Title = "Information about Exchange resource `"$resource`""
    if ( $Shared -and !$Calendar ) {
        $Title += " (Shared Mailbox)"
    } elseif ($Shared -and $Calendar ) {
        $Title += " (Shared Calendar)"
    } elseif ( $Equipment ) {
        $Title += " (Equipment Resource)"
    } elseif ( $Room ) {
        $Title += " (Room Resource)"
    }

    $To = (Get-Mailbox $Owner).PrimarySmtpAddress.ToString()

    $Body = @"
You have been identified as a resource owner / delegate for the
following Exchange resource:

    $resource`n

"@

    if ($Equipment -or $Room) {
        $Body += @"
This email is to inform you about the booking policy for this resource,
and how you can change it if the defaults do not suit the resource.
Currently, the resource will automatically accept booking requests
that do not conflict with other bookings, and will require your approval
if a request is made that conflicts with another booking.

If you would like to change this behavior, you may do so by using
Outlook Web Access (OWA).  Open Internet Explorer and navigate to the
following URL:`n
"@
    }

    if ( $Shared ) {
        $Body += @"
You may use either Outlook or Outlook Web Access (OWA) to access this 
resource.  If you would like to use OWA, open Internet Explorer and
navigate to the following URL:`n
"@
    }

    $Body += @"

    https://$($Fqdn)/owa/$($resource.PrimarySMTPAddress)`n

(Log in using your own e-ID and password.)`n

"@

    if ($Equipment -or $Room) {
        $Body += @"
Click on the Options link in the upper-right-hand corner, then click the
"Resource Settings" option in the left-hand column.  Most of the options
should be self-explanatory.  For instance, if you would like to alter
the settings of this resource such that no user can automatically book
it, and that every request must be approved, simply change both settings
that start with "These users can schedule automatically..." to "Select
users and groups" instead of "Everyone", and set "These users can submit
a request for manual approval..." to "Everyone".`n

"@
}

    $Body += @"
If you have any questions, please contact the JMU Computing HelpDesk at
helpdesk@jmu.edu, or by phone at 540-568-3555.
"@

    & "$cwd\Send-Email.ps1" -From $From -To $To -Bcc $Bcc -Subject $Title -Body $Body -SmtpServer $SmtpServer
    Write-Host "Sent message to $To for resource `"$resource`""
}
