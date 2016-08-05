################################################################################
#
# Copyright (c) 2009-2012 Seth Wright <wrightst@jmu.edu>
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

function Get-MailboxDatabaseStatistics {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [object]
            $Identity,

            [switch]
            $IncludeUserStatistics,

            [Parameter(Mandatory=$false)]
            [UInt64]
            $MaxDatabaseSizeInBytes=250GB
          )

    BEGIN {
        $i = 1
        $dstart = Get-Date
    }

    PROCESS {
        if ($Identity -is [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase]) {
            $db = $Identity
        } else {
            try {
                $db = Get-MailboxDatabase -Identity $Identity -Status -ErrorAction Stop
            } catch {
                throw
            }
        }

        $dend = Get-Date
        $dtotalSeconds = ($dend - $dstart).TotalSeconds
        $timePerDb = $dtotalSeconds / $i
        $dtimeLeft = $timePerDb * $i

        Write-Progress  -Activity "Gathering Database Statistics" `
                        -Status $db.Name `
                        -Id 1 -SecondsRemaining $dtimeLeft

        $dbInfo = New-Object PSObject -Property @{
                        Identity            = $db.Name
                        Mailboxes           = $null
                        EdbFileSize         = $null
                        AvailableSpace      = $null
                        CommitPercent       = $null
                        LastFullBackup      = $null
                        CILastModifiedTime  = @()
                        ServerStatus        = @()
                        MountedOnServer     = $db.MountedOnServer.Split(".")[0]
                    }

        if ($db.DatabaseSize -ne $null) {
            $dbInfo.EdbFileSize = $db.DatabaseSize
        }

        $dbInfo.LastFullBackup = $db.LastFullBackup

        for ($c = 0; $c -lt @($db.DatabaseCopies).Count; $c++) {
            $copy = $db.DatabaseCopies[$c]
            Write-Progress  -Activity "Retrieving statistics for copy $copy" `
                            -Status $copy.Identity `
                            -PercentComplete ($c/($db.DatabaseCopies).Count*100) `
                            -Id 2 -ParentId 1 `
                            -SecondsRemaining $dtimeLeft

            $status = Get-MailboxDatabaseCopyStatus -Identity $copy
            $serverStatus =
                New-Object PSObject -Property @{ DatabaseCopy = $copy; Status = $status.Status }
            $dbInfo.ServerStatus += $serverStatus
        }
        Write-Progress  -Activity "Retrieving statistics for copy $copy" `
                        -Status "Finished" `
                        -Id 2 -ParentId 1 `
                        -Completed

        $edbFilePath = $db.EdbFilePath.PathName
        $i = $edbFilePath.LastIndexOf('\')
        $ciPath = $edbFilePath.Remove($i)
        $ciPath = $ciPath.Replace(":", "$")
        $guid = $db.Identity.ObjectGuid.ToString()
        for ($c = 0; $c -lt @($db.Servers).Count; $c++) {
            $server = $db.Servers[$c]
            Write-Progress  -Activity "Retrieving LastWriteTime for Content Index directory" `
                            -Status $server `
                            -PercentComplete ($c/($db.Servers).Count*100) `
                            -Id 2 -ParentId 1 `
                            -SecondsRemaining $dtimeLeft

            $ciUncPath = "\\" + $server + "\" + $ciPath + "\CatalogData-" + $guid + "*"
            $ciUncPath = Resolve-Path $ciUncPath
            if ($ciUncPath -eq $null) {
                $ciLastModifiedTime =
                    New-Object PSObject -Property @{ Server = $server; LastModifiedTime = [DateTime]::MinValue }
            } else {
                $ciLastModifiedTime =
                    New-Object PSObject -Property @{ Server = $server; LastModifiedTime = (Get-Item $ciUncPath).LastWriteTime }
            }
            $dbInfo.CILastModifiedTime += $ciLastModifiedTime
        }
        Write-Progress  -Activity "Retrieving LastWriteTime for Content Index directory" `
                        -Status "Finished" `
                        -Id 2 -ParentId 1 `
                        -Completed

        $dbInfo.AvailableSpace = $db.AvailableNewMailboxSpace

        Write-Progress  -Activity "Gathering Database Statistics for $Identity" `
                        -Status "Retrieving mailbox count" `
                        -Id 1 -SecondsRemaining $dtimeLeft
        $dbUsers = @(Get-Mailbox -ResultSize Unlimited -Database $db)
        $dbInfo.Mailboxes = $dbUsers.Count

        if ($db.ProhibitSendReceiveQuota -eq "unlimited") {
            Write-Verbose "Database quota set to unlimited"
        } else {
            [UInt64]$totalDbUserQuota = 0
            $dbQuota = $db.ProhibitSendReceiveQuota.Value.ToBytes()
        }

        if ($IncludeUserStatistics) {
            $usersCount = $dbUsers.Count
            $j = 0
            $startTime = Get-Date
            foreach ($user in $dbUsers) {
                if ($db.ProhibitSendReceiveQuota -ne "unlimited") {
                    if ($user.UseDatabaseQuotaDefaults -eq $true) {
                        $totalDbUserQuota += $dbQuota
                    } else {
                        $userQuota = $user.ProhibitSendReceiveQuota
                        if ($userQuota.IsUnlimited -eq $false) {
                            $totalDbUserQuota += $userQuota.Value.ToBytes()
                        }
                    }
                }
                $j++
                $end = Get-Date
                $totalSeconds = ($end - $startTime).TotalSeconds
                $timePerUser = $totalSeconds / $j
                $timeLeft = $timePerUser * ($usersCount - $j)
                Write-Progress  -Activity "Gathering User Statistics for Database $($db.Name)" `
                                -Status $user `
                                -PercentComplete ($j/$usersCount*100) `
                                -Id 2 -ParentId 1 `
                                -SecondsRemaining $timeLeft
            }
            Write-Progress -Activity "Gathering User Statistics" -Status "Finished" -Id 2 -ParentId 1 -Completed

            if ($db.ProhibitSendReceiveQuota -eq "unlimited") {
                $dbInfo.CommitPercent = "unlimited"
            } else {
                $dbInfo.CommitPercent = [Math]::Round(($totalDbUserQuota/$MaxDatabaseSizeInBytes*100), 0)
            }
        }
        $i++
        Write-Progress -Activity "Gathering Database Statistics" -Status "Finished" -Id 1 -Completed

        Write-Verbose "Database Name: $($dbInfo.Identity)"
        Write-Verbose "Database Size: $($dbInfo.EdbFileSizeInGB)"
        Write-Verbose "Database Available Space: $($dbInfo.AvailableSpaceInMB)MB"
        Write-Verbose "Database Commit %: $($dbInfo.CommitPercent)"
        Write-Verbose "Database Last Full Backup: $($dbInfo.LastFullBackup)"
        Write-Verbose "Database Backup Status: $($dbInfo.BackupStatus)"
        Write-Verbose "Database Mounted on:  $($dbInfo.MountedOnServer)"

        return $dbInfo
    }
}
