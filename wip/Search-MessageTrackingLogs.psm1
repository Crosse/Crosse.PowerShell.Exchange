################################################################################
# 
# $Id$
# 
# DESCRIPTION:  Searches the Message Tracking logs on all hubs in the Exchange org.
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

param ( $Start=$(Get-Date).AddDays(-1),
        $End=$(Get-Date),
        $Sender=$null,
        $Recipients=$null,
        $EventId=$null,
        $MessageSubject=$null,
        $MessageId=$null) { }

$hubs = Get-ExchangeServer | Where-Object { $_.ServerRole -match 'Hub' }

$cmd = "Invoke-Command { `$hubs | Get-MessageTrackingLog -ResultSize Unlimited -Start `"$Start`""

if ($End -ne $null) {
    $cmd += " -End `"$End`""
}

if (![String]::IsNullOrEmpty($Sender)) {
    $cmd += " -Sender $Sender"
}

if (![String]::IsNullOrEmpty($Recipients)) {
    $cmd += " -Recipients $Recipients"
}

if (![String]::IsNullOrEmpty($EventId)) {
    $cmd += " -EventId $EventId"
}

if (![String]::IsNullOrEmpty($MessageSubject)) {
    $cmd += " -MessageSubject `"$MessageSubject`""
}

if (![String]::IsNullOrEmpty($MessageId)) {
    $cmd += " -MessageId `"$MessageId`""
}

$cmd += " }"

Invoke-Expression $cmd
