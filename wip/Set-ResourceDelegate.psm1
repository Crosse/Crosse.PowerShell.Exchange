################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
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

param ([string]$Identity, [string]$Delegate, [switch]$EmailDelegate=$true) {}

if ($Delegate -eq '' -or $Identity -eq '') {
    Write-Error "Please specify the Identity and Delegate"
    return
}

# Change these to suit your environment
$SmtpServer         = "it-exhub.ad.jmu.edu"
$From               = "it-exmaint@jmu.edu"
$Bcc                = "wrightst@jmu.edu, millerca@jmu.edu, najdziav@jmu.edu, eckardsl@jmu.edu"
$Fqdn               = "exchange.jmu.edu"
##################################
$cwd                = [System.IO.Path]::GetDirectoryName(($MyInvocation.MyCommand).Definition)
$DomainController   = (gc Env:\LOGONSERVER).Replace('\', '')

if ($DomainController -eq $null) { 
    Write-Warning "Could not determine the local computer's logon server!"
    return
}


$resource = Get-Mailbox $Identity
$objUser = Get-User $Delegate

if ($resource -eq $null) {
    Write-Error "Could not find Resource"
    return
}

Write-Host "Setting Permissions: "

# Grant Send-As rights to the owner:
$resource | Add-ADPermission -ExtendedRights "Send-As" -User $objUser.DistinguishedName `
            -DomainController $DomainController

# Give the owner Full Access to the resource:
$resource | Add-MailboxPermission -DomainController $DomainController `
            -AccessRights FullAccess -User $objUser.Identity

# Grant SendOnBehalfOf rights to the owner:
if ($objUser.RecipientType -match 'MailUser' -or $objUser.RecipientType -match 'UserMailbox') {
    $sobo = (Get-Mailbox -DomainController $DomainController -Identity $resource).GrantSendOnBehalfTo
    if ( !$sobo.Contains((Get-User $objUser).DistinguishedName) ) {
        $sobo.Add( (Get-User $objUser).DistinguishedName )
    }
    $resource | Set-Mailbox -DomainController $DomainController `
                -GrantSendOnBehalfTo $sobo
} else {
    Write-Output "Not setting Send-On-Behalf-Of rights, because $($objUser.SamAccountName) is a $($objUser.RecipientType)"
}

if ( ($resource.RecipientTypeDetails -eq 'RoomMailbox') -or ($resource.RecipientTypeDetails -eq 'EquipmentMailbox') ) {
    # Set the ResourceDelegates
    $resourceDelegates = (Get-CalendarProcessing -Identity $resource).ResourceDelegates
    if ( !($resourceDelegates.Contains((Get-User $objUser).DistinguishedName)) ) {
        $resourceDelegates.Add( (Get-User $objUser).DistinguishedName )
    }

    foreach ($i in 1..10) {
        $error.Clear()
        $resource | Set-CalendarProcessing -DomainController $DomainController `
                    -ResourceDelegates $resourceDelegates -ErrorAction SilentlyContinue
        if (![String]::IsNullOrEmpty($error[0])) {
            Write-Host -NoNewLine "."
            Start-Sleep $i
        } else {
            Write-Host "done."
            break
        }
    }
}

$resourceType = $resource.RecipientTypeDetails

if ($EmailDelegate -and $objUser.RecipientTypeDetails -ne "User" ) {
    $Title = "Information about Exchange resource `"$resource`""
    if ( $resourceType -eq 'SharedMailbox' ) {
        $Title += " (Shared Mailbox)"
    } elseif ( $resourceType -eq 'EquipmentMailbox' ) {
        $Title += " (Equipment Resource)"
    } elseif ( $resourceType -eq 'RoomMailbox' ) {
        $Title += " (Room Resource)"
    }

    $To = (Get-Recipient $objUser.Identity).PrimarySmtpAddress.ToString()

    $Body = @"
You have been identified as a resource owner / delegate for the
following Exchange resource:

    $resource`n

"@

    if (($resourceType -eq 'RoomMailbox') -or ($resourceType -eq 'EquipmentMailbox') ) {
        $Body += @"
This email is to inform you about the booking policy for this resource,
and how you can change it if the defaults do not suit the resource.
By default, the resource will automatically accept booking requests
that do not conflict with other bookings, and will require your approval
if a request is made that conflicts with another booking.

If you would like to change this behavior, you may do so by using
Outlook Web Access (OWA).  Open Internet Explorer and navigate to the
following URL:`n
"@
    }

    if ( $resourceType -eq 'SharedMailbox' ) {
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

    if ( ($resourceType -eq 'EquipmentMailbox') -or ($resourceType -eq 'RoomMailbox') ) {
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
    Write-Output "Sent message to $To for resource `"$resource`""
}

