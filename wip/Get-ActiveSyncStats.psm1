$users = Get-CASMailbox -ResultSize Unlimited -Filter { HasActiveSyncDevicePartnership -eq $true -and -not DisplayName -like "CAS_{*" } | Get-Mailbox -ResultSize Unlimited

foreach ($user in $users) { 
    Get-ActiveSyncDeviceStatistics -Mailbox $user -ErrorAction SilentlyContinue 
} | Select-Object Identity,FirstSyncTime,LastSyncSuccess,LastSyncAttemptTime,Devicetype,DeviceModel,DeviceUserAgent,DevicePhoneNumber

