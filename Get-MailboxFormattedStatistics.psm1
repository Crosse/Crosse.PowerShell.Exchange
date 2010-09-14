################################################################################
# 
# $Id$
# 
# DESCRIPTION:  Gets mailbox data with the sizes formatted in MB.
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

param ($User=$null, $inputObject=$null)

# This section executes only once, before the pipeline.
BEGIN {
    if ($inputObject) {
        Write-Output $inputObject | &($MyInvocation.InvocationName)
        break
    }

} # end 'BEGIN{}'

# This section executes for each object in the pipeline.
PROCESS {
    if ($_) { $User = $_ }

    if ($User -eq $null) {
        Write-Warning "No user specified; generating statistics for all Mailboxes"
    } else {
        $User = Get-Mailbox $User.ToString() -ErrorAction SilentlyContinue

        if ($User -eq $null) {
            Write-Host "User $User does not have a mailbox."
            return
        }
    }


    Get-Mailbox $User -ResultSize Unlimited | 
        Get-MailboxStatistics | 
        Select-Object DisplayName, `
            @{Name="QuotaStatus"; `
                Expression={$_.StorageLimitStatus} }, `
            @{Name="ItemSize(MB)"; `
                Expression={$_.TotalItemSize.Value.ToMB() } }, `
            @{Name="DeletedItemSize(MB)"; `
                Expression={$_.TotalDeletedItemSize.Value.ToMB() } }, `
            ItemCount, DeletedItemCount
}
