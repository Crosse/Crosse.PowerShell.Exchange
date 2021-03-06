################################################################################
#
# Copyright (c) 2013 Seth Wright <seth@crosse.org>
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
    Retrieves SMTP logs from SmtpSend and SmtpReceive log files.

    .DESCRIPTION
    Retrieves SMTP logs from SmtpSend and SmtpReceive log files generated by
    an Exchange Hub or Edge Transport server.

    .INPUTS
    None.  Get-SmtpLogs does not take any input from the pipeline.

    .OUTPUTS
    Things.

    .EXAMPLE
    C:\PS> $results = Get-SmtpLogs -ClientIpAddress 192.168.1.25 -LogFileBasePath 'C:\Program Files\Microsoft\Exchange Server\V14\TransportRoles\Logs\ProtocolLog\SmtpReceive'
    C:\PS> $results
    SessionId       : 08CFDCCAF9C8DAED
    ClientPort      : 1024
    ClientIpAddress : 192.168.1.25
    ConnectorId     : hub1\DEFAULT
    ServerIpAddress : 192.168.1.20
    Conversation    : {@{SequenceNumber=0; Context=; Data=; Event=+}, @{Sequenc
                      eNumber=1; Context=Set Session Permissions;Data=SMTPSubmi
                      t SMTPAcceptAnySender SMTPAcceptAuthoritativeDomainSender
                       AcceptRoutingHeaders; Event=*}, @{SequenceNumber=2; Cont
                      ext=; Data=220 hub1.contoso.local Microsoft ESMTP MAIL Se
                      rvice ready at Thu, 28 Feb 2013 15:05:10 -0500; Event=>},
                       @{SequenceNumber=3; Context=; Data=HELO server2.contoso.
                      local; Event=<}...}
    ServerPort      : 25
    DateTime        : 2013-02-28 20:05:10

    C:\PS> $results.Conversation | ft -AutoSize Event,Data
    Event Data
    ----- ----
    +
    *     SMTPSubmit SMTPAcceptAnySender SMTPAcceptAuthoritativeDomainSender AcceptRoutingHeaders
    >     220 hub1.contoso.local Microsoft ESMTP MAIL Service ready at Thu, 28 Feb 2013 15:05:10 -0500
    <     HELO server2.contoso.local
    >     250 hub1.contoso.local Hello [134.126.31.254]
    <     MAIL FROM: <user1@contoso.local>
    *     08CFDCCAF9C8DAED;2013-02-28T20:05:10.994Z;1
    >     250 2.1.0 Sender OK
    <     RCPT TO: <user2@contoso.local>
    >     250 2.1.5 Recipient OK
    <     DATA
    >     354 Start mail input; end with <CRLF>.<CRLF>
    *     Tarpit for '0.00:00:00.561' due to 'DelayedAck'
    >     250 2.6.0 <512F7F80.001.00206B7E4696.user1@contoso.local> [InternalId=77845278] Queued mail for delivery
    <     QUIT
    >     221 2.0.0 Service closing transmission channel
    -

#>
################################################################################
function Get-SmtpLogs {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ParameterSetName="Client")]
            [Parameter(Mandatory=$false,
                ParameterSetName="Server")]
            [string]
            # Specifies the client IP address to search for.
            $ClientIpAddress,

            [Parameter(Mandatory=$true,
                ParameterSetName="Server")]
            [Parameter(Mandatory=$false,
                ParameterSetName="Client")]
            [string]
            # Specifies the server IP address to search for.
            $ServerIpAddress,

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
            # Specifies that the LogFileBasePath should be recursed to find all files that meet the Start and End parameters.
            [switch]
            $Recurse,

            [Parameter(Mandatory=$false)]
            # Specifies the maximum number of results to display.  By default, all results are displayed.
            [Nullable[int]]
            $ResultSize,

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

    $logfiles = @()
    foreach ($path in $LogFileBasePath) {
        if ((Test-Path $path) -eq $false) {
            Write-Warning "Path not found: $path"
        } else {
            # Get a list of all files modified after $Start
            $files = Get-ChildItem $path -File -Filter *.LOG -Recurse:$Recurse | Where-Object {
                $_.LastWriteTime -gt $Start -and $_.LastWriteTime -lt $End
            }
            $logfiles += $files
        }
    }
    if ($logfiles.Count -eq 0) {
        Write-Error "No matching log files were found."
        return
    }

    Write-Verbose "Found $($logfiles.Count) total log files."
    $fromFiles += ($logfiles | Foreach-Object { $_.FullName }) -join ",`n    "

    if ([String]::IsNullOrEmpty($ClientIpAddress) -eq $false) {
        $clientAddr = "ClientIpAddress = '$ClientIpAddress' AND"
    }

    if ([String]::IsNullOrEmpty($ServerIpAddress) -eq $false) {
        $serverAddr = "ServerIpAddress = '$ServerIpAddress' AND"
    }

    $formattedStartTime = $Start.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
    $formattedEndTime = $End.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')

    $query = @"
SELECT
    TO_TIMESTAMP(date-time, 'yyyy-MM-dd?hh:mm:ss.ll?') AS DateTime,
    [connector-id] AS ConnectorId,
    [session-id] AS SessionId,
    [sequence-number] AS SequenceNumber,
    EXTRACT_TOKEN(local-endpoint, 0, ':') AS ServerIpAddress,
    EXTRACT_TOKEN(local-endpoint, 1, ':') AS ServerPort,
    EXTRACT_TOKEN(remote-endpoint, 0, ':') AS ClientIpAddress,
    EXTRACT_TOKEN(remote-endpoint, 1, ':') AS ClientPort,
    [event] AS Event,
    [data] AS Data,
    [context] AS Context
FROM
    $fromFiles
WHERE
    $clientAddr
    $serverAddr
    DateTime BETWEEN
        TO_TIMESTAMP('$formattedStartTime', 'yyyy-MM-dd HH:mm:ss') AND
        TO_TIMESTAMP('$formattedEndTime', 'yyyy-MM-dd HH:mm:ss')
ORDER BY
   DateTime,SessionId,SequenceNumber
"@

    if ($WhatIf) {
        Write-Host "WhatIf:"
        Write-Host $query
    } else {
        $header = "date-time,connector-id,session-id,sequence-number,local-endpoint,remote-endpoint,event,data,context"
        $headerFile = [IO.Path]::GetTempFileName()
        $header | Out-File -Encoding ASCII -FilePath $headerFile
        Write-Verbose "Running query..."
        $begin = Get-Date
        $results = & $LogParserLocation -stats:OFF -i:CSV -nSkipLines:1 -comment "#" -headerRow:OFF -iHeaderFile $headerFile -o:CSV "$query"
        Write-Verbose "Query took $(((Get-Date) - $begin).TotalSeconds) seconds."
        if ($results -eq $null) {
            Write-Warning "Query produced no results."
        } else {
            $results = $results | ConvertFrom-Csv
            $hash = @{}
            foreach ($result in $results) {
                $id = $result.ServerIpAddress + ":" + $result.SessionId + ":" + $result.ClientIpAddress + ":" + $result.ClientPort
                if ($hash[$id] -eq $null) {
                    $hash[$id] = New-Object PSObject -Property @{
                        DateTime = $result.DateTime
                        ConnectorId = $result.ConnectorId
                        SessionId = $result.SessionId
                        ServerIpAddress = $result.ServerIpAddress
                        ServerPort = $result.ServerPort
                        ClientIpAddress = $result.ClientIpAddress
                        ClientPort = $result.ClientPort
                        Conversation = @()
                    }
                }
                $convo = New-Object PSObject -Property @{
                    SequenceNumber = $result.SequenceNumber
                    Event = $result.Event
                    Data = $result.Data
                    Context = $result.Context
                }
                $hash[$id].Conversation += $convo
            }
            if ($ResultSize -eq $null) {
                $hash.Values | Sort-Object DateTime,SessionId
            } else {
                $hash.Values | Sort-Object DateTime,SessionId | Select-Object -First $ResultSize
            }
        }
    }
}
