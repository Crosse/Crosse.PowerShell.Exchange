################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  Returns the best database in which to create a new mailbox.
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

param ([string]$Server="localhost", [switch]$Single=$true)

##################################

$srv = Get-ExchangeServer $Server
if ($srv -eq $null) {
    Write-Error "Could not find Exchange Server $Server"
    return
}

$DomainController = (gc Env:\LOGONSERVER).Replace('\', '')
if ($DomainController -eq $null) { 
    Write-Warning "Could not determine the local computer's logon server!"
    return
}

$databases = Get-MailboxDatabase -Server $srv -Status | 
Where { $_.Mounted -eq $True -and $_.Name -match "^(SG|DB)" }
if ($databases -eq $null) {
    Write-Error "Could not enumerate databases on server $Server"
    return
}

$i = 1
$dbs = New-Object System.Collections.Hashtable
foreach ($database in $databases) {
    $percent = $([int]($i/$databases.Count*100))
    Write-Progress -Activity "Processing Mailbox Databases" `
        -Status "$percent% Complete" `
        -PercentComplete $percent -CurrentOperation "Verifying $($database.Identity)"
    $i++

    $mailboxCount = @((Get-Mailbox -Database $database)).Count
    if ($? -eq $False) {
        Write-Error "Error processing database $database"
        return
    }

#    $maxUsers = 200GB / (Get-MailboxDatabase $database).ProhibitSendReceiveQuota.Value.ToBytes()

#    if ($mailboxCount -le $maxUsers) {
# Normally we'd not add this database
# if the mailboxCount was greater than the maximum
# number of users allowed for the database,
# but we're fudging it for a while.
#    }
    $dbs.Add($database.Identity.ToString(), $mailboxCount)
}

Write-Progress -Activity "Processing Mailbox Databases" -Status "100% Complete" -Complete:$true

if ($Single) {
    $candidate = $null
    foreach ($db in $dbs.Keys) {
        if ($candidate -eq $null) {
            $candidate = $db
        } else {
            if ($dbs[$db] -lt $dbs[$candidate]) {
                $candidate = $db
            }
        }
    }
    $candidate
} else {
    $dbs
}
