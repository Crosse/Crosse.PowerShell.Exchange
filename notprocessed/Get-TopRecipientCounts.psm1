################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  Gets the top senders, based on total number of recipients
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

param ( $StartDate=(Get-Date).AddDays(-1),
        $ResultsToReturn=10) { }

$hubs = Get-ExchangeServer | ? { $_.ServerRole -match "HubTransport" }
$recipientCounts = New-Object System.Collections.Hashtable

$msgs = $hubs | Get-MessageTrackingLog -Start $StartDate -EventId RECEIVE `
                    -ResultSize Unlimited | ? { 
                        $_.Source -match 'STOREDRIVER' }

foreach ($msg in $msgs) { 
    $sender = $msg.Sender
        if ($recipientCounts.Contains($sender)) {
            $recipientCounts[$sender] += $msg.RecipientCount
        } else { 
            $recipientCounts.Add($sender, 1)
        }
}

$recipientCounts.GetEnumerator() | Sort Value -Descending | Select -First $ResultsToReturn
