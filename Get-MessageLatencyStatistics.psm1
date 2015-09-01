function Get-MessageLatencyStatistics {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$false,
                ValueFromPipeline=$true)]
            [Microsoft.Exchange.Management.TransportLogSearchTasks.MessageTrackingEvent[]]
            $Events,

            [switch]
            $IncludeMessages
        )

    BEGIN {
        $totalMessages = 0
        $delayHash = @{}
        $breakpoints = @(
                [TimeSpan]::FromSeconds(10)
                [TimeSpan]::FromSeconds(30)
                [TimeSpan]::FromSeconds(60)
                [TimeSpan]::FromSeconds(180)
                [TimeSpan]::FromSeconds(300)
                [TimeSpan]::FromSeconds(600)
                )

        foreach ($bp in $breakpoints) { $delayHash[$bp] = @(0, @()) }
    }

    PROCESS {
        foreach ($event in $Events) {
            if ($event.EventID -ne "SEND") { continue }
            $totalMessages += 1

            foreach ($bp in $breakpoints) {
                if ($event.MessageLatency -gt $bp) {
                    $delayHash[$bp][0] += 1
                    if ($IncludeMessages) {
                        $delayHash[$bp][1] += $event
                    }
                }
            }
        }

    }

    END {
        Write-Verbose "Found $totalMessages messages"

        foreach ($bp in $breakpoints) {
            $numDelayed = $delayHash[$bp][0]
            $pctDelayed = [Math]::Round(($numDelayed / $totalMessages) * 100, 2)
            New-Object PSObject -Property @{
                Messages = $delayHash[$bp][1]
                DelayedAtLeast = $bp
                DelayedPercent = $pctDelayed
                DelayedMessages = $numDelayed
            }
        }
    }
}
