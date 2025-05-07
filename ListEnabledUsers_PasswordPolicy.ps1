# Get all users with key properties
$allUsers = Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName,AccountEnabled,PasswordPolicies,LastPasswordChangeDateTime"

# Get sign-in logs (limited to last 7 days unless retention is extended)
$signInLogs = Get-MgAuditLogSignIn -All

# Get audit logs to check for disable events
$auditLogs = Get-MgAuditLogDirectoryAudit -All | Where-Object {
    $_.ActivityDisplayName -eq "Update user" -and
    $_.TargetResources.ModifiedProperties |
        Where-Object { $_.DisplayName -eq "AccountEnabled" -and $_.NewValue -eq "false" }
}

# Process user data
$results = foreach ($user in $allUsers) {
    $lastSignIn = ($signInLogs | Where-Object { $_.UserId -eq $user.Id } |
        Sort-Object CreatedDateTime -Descending | Select-Object -First 1).CreatedDateTime

    $disableEvent = ($auditLogs | Where-Object {
        $_.TargetResources[0].Id -eq $user.Id
    } | Sort-Object ActivityDateTime -Descending | Select-Object -First 1).ActivityDateTime

    [PSCustomObject]@{
        DisplayName               = $user.DisplayName
        UserPrincipalName         = $user.UserPrincipalName
        AccountEnabled            = $user.AccountEnabled
        PasswordPolicies          = $user.PasswordPolicies
        LastPasswordChangeDate    = $user.LastPasswordChangeDateTime
        LastSignIn                = $lastSignIn
        DateDisabled              = if (-not $user.AccountEnabled) { $disableEvent } else { $null }
    }
}

# Export to CSV
$results | Export-Csv -Path "S:\Tech Team\Benjamin\Scripts\EntraUsersReport.csv" -NoTypeInformation
