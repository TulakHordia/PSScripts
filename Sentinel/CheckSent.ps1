# Define the output file
$outputFile = "C:\Scripts\SentinelOne_NotInstalled.txt"

# Get a list of all computers in the domain
$computers = Get-ADComputer -Filter * -Property Name | Select-Object -ExpandProperty Name

# Initialize an array to store computers without SentinelOne
$computersWithoutSentinel = @()

# SentinelOne display name (adjust if necessary)
$sentinelOneName = "SentinelOne"

foreach ($computer in $computers) {
    Write-Host "Checking $computer..." -ForegroundColor Cyan
    try {
        # Check if the computer is online
        if (Test-Connection -ComputerName $computer -Count 1 -Quiet) {
            # Query installed software
            $software = Get-WmiObject -Class Win32_Product -ComputerName $computer -ErrorAction Stop |
                        Where-Object { $_.Name -like "*$sentinelOneName*" }
            
            if (-not $software) {
                Write-Host "$computer does not have SentinelOne installed." -ForegroundColor Red
                $computersWithoutSentinel += $computer
            } else {
                Write-Host "$computer has SentinelOne installed." -ForegroundColor Green
            }
        } else {
            Write-Host "$computer is offline or unreachable." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error checking $computer: $_" -ForegroundColor Magenta
    }
}

# Output the results to a text file
$computersWithoutSentinel | Out-File -FilePath $outputFile -Encoding UTF8

Write-Host "Scan complete. Results saved to $outputFile." -ForegroundColor Green
