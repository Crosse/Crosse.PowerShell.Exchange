################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  Expands nested distribution groups
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

function Expand-DistributionGroup {
    [CmdletBinding(SupportsShouldProcess=$true,
        ConfirmImpact="High")]

    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [Alias("Name")]
            # The group to expand
            $Identity
          )

    # This section executes only once, before the pipeline.
    BEGIN {
        # Nothing to see here, move along.
    }

    # This section executes for each value in the pipeline.
    PROCESS {
        $group = Get-DistributionGroup $Identity
        if ($group -eq $null) {
            Write-Error "Group $Identity does not exist"
            return
        }

        $retval =  New-Object System.Collections.ArrayList
        $members = Get-DistributionGroupMember $group.Name

        foreach ($member in $members) {
            if ($member.RecipientType -match "Group") {
                $retval.AddRange($(Expand-DistributionGroup $member.Name)) | Out-Null
            } else {
                $retval.Add($member) | Out-Null
            }
        }

        $retval | Sort-Object -Unique
    }

    # This section executes only once, after the pipeline.
    END {
        # Nothing to see here, either.
    }
}
