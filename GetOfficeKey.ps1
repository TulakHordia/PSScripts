function Get-OfficeProductKey {
    $OfficePath = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Office" -Recurse |
                  Where-Object { $_.Name -match "Registration" -and $_.Property -contains "DigitalProductID" }
    foreach ($Path in $OfficePath) {
        $DigitalProductID = (Get-ItemProperty -Path $Path.PSPath).DigitalProductID
        $Key = ConvertFrom-DigitalProductID $DigitalProductID
        if ($Key) { $Key }
    }
}

function ConvertFrom-DigitalProductID {
    param ([byte[]]$DigitalProductID)
    $KeyChars = "BCDFGHJKMPQRTVWXY2346789"
    $ProductKey = ""
    for ($i = 24; $i -ge 0; $i--) {
        $Current = 0
        for ($j = 14; $j -ge 0; $j--) {
            $Current = $Current * 256 -bxor $DigitalProductID[$j]
            $DigitalProductID[$j] = [math]::Floor($Current / 24)
            $Current = $Current % 24
        }
        $ProductKey = $KeyChars[$Current] + $ProductKey
    }
    return $ProductKey -replace ".{5}", '$&-'
}

Get-OfficeProductKey
