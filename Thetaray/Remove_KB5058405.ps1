# Block and Remove Specific KB Update via Intune
$kbID = "KB5058405"  # Change this to your target KB

# Uninstall the KB if installed
$hotfix = Get-HotFix | Where-Object { $_.HotFixID -eq $kbID }
if ($hotfix) {
    Write-Output "Uninstalling $($kbID)..."
    try {
        Start-Process "wusa.exe" "/uninstall /kb:$($kbID.Replace('KB','')) /quiet /norestart" -Wait
        Write-Output "$kbID uninstall triggered."
    } catch {
        Write-Output "Failed to uninstall $($kbID): $_"
    }
} else {
    Write-Output "$kbID is not currently installed."
}

# Block the KB from reinstalling
try {
    $Session = New-Object -ComObject Microsoft.Update.Session
    $Searcher = $Session.CreateUpdateSearcher()
    $SearchResult = $Searcher.Search("IsInstalled=0")

    foreach ($update in $SearchResult.Updates) {
        if ($update.Title -like "*$kbID*") {
            Write-Output "Hiding update: $($update.Title)"
            $update.IsHidden = $true
        }
    }
} catch {
    Write-Output "Error while hiding update: $_"
}
