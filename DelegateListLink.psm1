function Get-MailboxDelegateListLinkAttribute {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [object]
            $Identity
          )

    PROCESS {
        if ($Identity -is [Microsoft.Exchange.Data.Directory.Management.Mailbox]) {
            $mbox = $Identity
        } else {
            $mbox = Get-Mailbox -Identity $Identity -ErrorAction Stop
        }
        Write-Verbose "Mailbox DisplayName: $($mbox.DisplayName)"

        $attr = Get-ADAttribute -Identity $mbox.DistinguishedName `
                                -Attributes msExchDelegateListLink
        $delegates = @($attr.msExchDelegateListLink)
        Write-Verbose "Found $($delegates.Count) delegates"

        $users = @()
        foreach ($dn in $delegates) {
            if ([String]::IsNullOrEmpty($dn)) { continue }
            Write-Verbose "Looking up DN `"$dn`""
            $users += Get-User -Identity $dn
        }

        New-Object PSObject -Property @{
            Name                = $mbox.DisplayName
            DelegateListLinks   = $users
        }
    }
}

function Update-MailboxDelegateListLinkAttribute {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [object]
            $Identity,

#            [Parameter(Mandatory=$true,
#                ParameterSetName="Add")]
#            [ValidateNotNullOrEmpty()]
#            [object[]]
#            $Add,

            [Parameter(Mandatory=$true,
                ParameterSetName="Remove")]
            [ValidateNotNullOrEmpty()]
            [object[]]
            $Remove
          )

    if ($Identity -is [Microsoft.Exchange.Data.Directory.Management.Mailbox]) {
        $mbox = $Identity
    } else {
        $mbox = Get-User -Identity $Identity -ErrorAction Stop
    }
    Write-Verbose "Mailbox DisplayName: $($mbox.DisplayName)"

    $attr = Get-ADAttribute -Identity $mbox.DistinguishedName `
                            -Attributes msExchDelegateListLink
    $delegates = @($attr.msExchDelegateListLink)

    $users = New-Object System.Collections.ArrayList
    foreach ($user in $Remove) {
        $dn = (Get-User -Identity $user -ErrorAction Stop).DistinguishedName
        $null = $users.Add($dn)
    }

    $dnList = New-Object System.Collections.ArrayList
    foreach ($user in $users) {
        $found = $false
        foreach ($dn in $delegates) {
            if ($user -eq $dn) {
                $found = $true
                break
            }
        }
        if (!$found) {
            Write-Warning "$user is not in the delegate list"
        }
        $null = $dnList.Add($user)
    }

    foreach ($dn in $dnList) {
        $null = $users.Remove($dn)
    }

    if ($users.Count -eq 0) {
        $users = $null
    }

    Write-Verbose "Calling Set-ADAttribute"
    $null = Set-ADAttribute -Identity $Identity `
                    -Attribute msExchDelegateListLink `
                    -Value $users -Confirm:$false

    Get-ADAttribute -Identity $Identity -Attribute msExchDelegateListLink
}
