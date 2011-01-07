################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  Rebalances databases according to their activation preference.
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

function Rebalance-MailboxDatabases {
    [CmdletBinding(SupportsShouldProcess=$true,
        ConfirmImpact="High")]

    param ()

    $caption = "Confirm"
    $verboseDescription = "Verbose Description"
    foreach ($db in Get-MailboxDatabase) {
        $warning = "Are you sure you want to perform this action?"
        $warning += "`n"

        $pref = $db.ActivationPreference | ? { $_.Value -eq 1 }
        $server = $pref.Key.Name
        if ($db.Server -notmatch $server) {
            $warning += "Moving mailbox database `"$db`" from server `"$($db.Server)`" to server `"$server`"."
            if ($PSCmdlet.ShouldProcess($verboseDescription, $warning, $caption)) {
                Move-ActiveMailboxDatabase -Identity $db `
                                            -ActivateOnServer $server `
                                            -MoveComment "Database Rebalance" `
                                            -Confirm:$false
            }
            Write-Host ""
        }
    }
}
