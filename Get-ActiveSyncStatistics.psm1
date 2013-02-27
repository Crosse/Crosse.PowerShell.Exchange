################################################################################
#
# Copyright (c) 201# Seth Wright <seth@crosse.org>
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

################################################################################
<#
    .SYNOPSIS
    Retrieves ActiveSync statistics from IIS log files.

    .DESCRIPTION
    Retrieves ActiveSync statistics from the IIS log files produced by a Client
    Access Server.

    .INPUTS
    None.  Get-ActiveSyncStatistics does not take any input from the pipeline.

    .OUTPUTS
    Things.

    .EXAMPLE
    C:\PS> Get-ActiveSyncStatistics -LogFileBasePath "C:\inetpub\logs\LogFiles\W3SVC1" -User steve
    UserName        : steve
    DeviceId        : 321F17A7A0AFBFFE6D289B2D5391E
    DeviceType      : WP8
    DeviceUserAgent :
    Ping            : 0
    Sync            : 50
    FolderSync      : 0
    GetItemEstimate : 0
    MeetingResponse : 0
    Search          : 0
    SendMail        : 0
    MoveItems       : 0
    Settings        : 0
    GetAttachment   : 0
    Provision       : 0
    SmartReply      : 0
    SmartForward    : 0
    ItemOperations  : 0
    FolderCreate    : 0
    FolderDelete    : 0
    TotalHits       : 50
#>
################################################################################
function Get-ActiveSyncStatistics {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ParameterSetName="User")]
            [string]
            # Specifies the username (without the domain) to search for.
            $User,

            [Parameter(Mandatory=$false,
                ParameterSetName="Device")]
            # Specifies the device id to search for.
            [string]
            $Device,

            [Parameter(Mandatory=$false)]
            # Specifies the start date and time to return details.  The date and time must be specified in local time and will be converted appropriately.
            [DateTime]
            $Start = (Get-Date).AddDays(-1),

            [Parameter(Mandatory=$false)]
            # Specifies the end date and time to return details.  The date and time must be specified in local time and will be converted appropriately.
            [DateTime]
            $End = (Get-Date),

            [Parameter(Mandatory=$true)]
            # Specifies the base path where the client access server log files can be found.  If there are multiple locations, separate them with a comma (i.e., specify an array of log file locations).
            [string[]]
            $LogFileBasePath,

            [Parameter(Mandatory=$false)]
            # Specifies the maximum number of results to display.  By default, all results are displayed.
            [Nullable[int]]
            $ResultSize,

            [Parameter(Mandatory=$false)]
            # Specifies that the results should be sorted by hits, returning the devices with the highest number of hits.
            [switch]
            $Descending,

            [Parameter(Mandatory=$false)]
            # Specifies the minumum number of "hits" a device must have generated in order for it to be included in the results.  A "hit" equates to any ActiveSync command.
            [Nullable[int]]
            $MinimumHits,

            [Parameter(Mandatory=$false)]
            # Specifies the location of the LogParser.exe executable.  The default location is "C:\Program Files (x86)\Log Parser 2.2\logparser.exe".
            [string]
            $LogParserLocation = "C:\Program Files (x86)\Log Parser 2.2\logparser.exe",

            [Parameter(Mandatory=$false)]
            # Specifies that the query should be shown, but not actually run.
            [switch]
            $WhatIf
        )

    $version = & $LogParserLocation | Select-String -Pattern 'Microsoft.*Log Parser Version' -Quiet
    if ($version -eq $false) {
        Write-Error "Cannot find LogParser.exe"
        exit
    }

    if ($ResultSize -ne $null) {
        $top = "TOP $ResultSize"
    }

    $logfiles = @()
    foreach ($path in $LogFileBasePath) {
        if ((Test-Path $path) -eq $false) {
            Write-Warning "Path not found: $path"
        } else {
            # Get a list of all files modified after $Start
            $files = Get-ChildItem $path | Where-Object { 
                $_.LastWriteTime -gt $Start -and $_.LastWriteTime -lt $End
            }
            $logfiles += $files
        }
    }
    Write-Verbose "Found $($logfiles.Count) total log files."
    $fromFiles += ($logfiles | % { $_.FullName }) -join ",`n    "

    if ([String]::IsNullOrEmpty($User) -eq $false) {
        $whereUser = "`n    AND UserName  = '$User'"
    }

    if ([String]::IsNullOrEmpty($Device) -eq $false) {
        $whereDevice = "`n    AND DeviceId = '$Device'"
    }

    if ($MinimumHits -ne $null) {
        $havingHits = "HAVING TotalHits > $MinimumHits"
    }

    if ($Descending) {
        $orderBy = "ORDER BY TotalHits DESC"
    }

    $formattedStartTime = $Start.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
    $formattedEndTime = $End.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')

    $query = @"
SELECT $top
    TO_LOWERCASE(
            EXTRACT_PREFIX(
                EXTRACT_SUFFIX(cs-username, 0, '\\'),
                0, '@')
            ) AS UserName,
    EXTRACT_VALUE(cs-uri-query, 'DeviceId', '&') AS DeviceId,
    EXTRACT_VALUE(cs-uri-query, 'DeviceType', '&') AS DeviceType,
    cs(User-Agent) AS DeviceUserAgent,
    SUM(CASE Command WHEN 'Ping' THEN 1 ELSE 0 END) AS Ping,
    SUM(CASE Command WHEN 'Sync' THEN 1 ELSE 0 END) AS Sync,
    SUM(CASE Command WHEN 'FolderSync' THEN 1 ELSE 0 END) AS FolderSync,
    SUM(CASE Command WHEN 'GetItemEstimate' THEN 1 ELSE 0 END) AS GetItemEstimate,
    SUM(CASE Command WHEN 'MeetingResponse' THEN 1 ELSE 0 END) AS MeetingResponse,
    SUM(CASE Command WHEN 'Search' THEN 1 ELSE 0 END) AS Search,
    SUM(CASE Command WHEN 'SendMail' THEN 1 ELSE 0 END) AS SendMail,
    SUM(CASE Command WHEN 'MoveItems' THEN 1 ELSE 0 END) AS MoveItems,
    SUM(CASE Command WHEN 'Settings' THEN 1 ELSE 0 END) AS Settings,
    SUM(CASE Command WHEN 'GetAttachment' THEN 1 ELSE 0 END) AS GetAttachment,
    SUM(CASE Command WHEN 'Provision' THEN 1 ELSE 0 END) AS Provision,
    SUM(CASE Command WHEN 'SmartReply' THEN 1 ELSE 0 END) AS SmartReply,
    SUM(CASE Command WHEN 'SmartForward' THEN 1 ELSE 0 END) AS SmartForward,
    SUM(CASE Command WHEN 'ItemOperations' THEN 1 ELSE 0 END) AS ItemOperations,
    SUM(CASE Command WHEN 'FolderCreate' THEN 1 ELSE 0 END) AS FolderCreate,
    SUM(CASE Command WHEN 'FolderDelete' THEN 1 ELSE 0 END) AS FolderDelete,
    COUNT(*) AS TotalHits
USING
    EXTRACT_VALUE(cs-uri-query, 'Cmd', '&') AS Command,
    TO_TIMESTAMP(date, time) AS DateTime
FROM
    $fromFiles
WHERE
    cs-uri-stem LIKE '/Microsoft-Server-ActiveSync%'
    AND cs-method = 'POST' $whereUser $whereDevice
    AND DateTime BETWEEN
        TO_TIMESTAMP('$formattedStartTime', 'yyyy-MM-dd HH:mm:ss') AND
        TO_TIMESTAMP('$formattedEndTime', 'yyyy-MM-dd HH:mm:ss')
GROUP BY Username, DeviceId, DeviceType, DeviceUserAgent
$havingHits
$orderBy
"@

    if ($WhatIf) {
        Write-Host "WhatIf:"
        Write-Host $query
    } else {
        Write-Verbose "Running query..."
        $begin = Get-Date
        $results = & $LogParserLocation -stats:OFF -o:CSV "$query"
        Write-Verbose "Query took $(((Get-Date) - $begin).TotalSeconds) seconds."
        if ($results -eq $null) {
            Write-Warning "Query produced no results."
        } else {
            $results | ConvertFrom-Csv
        }
    }
}
