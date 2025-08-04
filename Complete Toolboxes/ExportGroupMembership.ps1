Add-Type -AssemblyName System.Windows.Forms

# יצירת טופס
$form = New-Object System.Windows.Forms.Form
$form.Text = "Group Membership Exporter"
$form.Width = 480
$form.Height = 260
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Label - שם משתמש
$labelUser = New-Object System.Windows.Forms.Label
$labelUser.Text = ":(הכנס שם משתמש (שם התחברות או שם מלא"
$labelUser.Top = 20
$labelUser.Left = 20
$labelUser.Width = 420
$form.Controls.Add($labelUser)

# TextBox - שם משתמש
$textBoxUser = New-Object System.Windows.Forms.TextBox
$textBoxUser.Top = 45
$textBoxUser.Left = 20
$textBoxUser.Width = 420
$form.Controls.Add($textBoxUser)

# Label - שם מבצע
$labelOperator = New-Object System.Windows.Forms.Label
$labelOperator.Text = ":הכנס את שם המבצע"
$labelOperator.Top = 80
$labelOperator.Left = 20
$labelOperator.Width = 420
$form.Controls.Add($labelOperator)

# TextBox - שם מבצע
$textBoxOperator = New-Object System.Windows.Forms.TextBox
$textBoxOperator.Top = 105
$textBoxOperator.Left = 20
$textBoxOperator.Width = 420
$form.Controls.Add($textBoxOperator)

# כפתור הרצה
$button = New-Object System.Windows.Forms.Button
$button.Text = "משוך קבוצות ושמור"
$button.Top = 150
$button.Left = 20
$button.Width = 200
$button.Height = 35
$form.Controls.Add($button)

# אופציונלי: הפעלת פעולה בלחיצה על Enter
$form.AcceptButton = $button

# פעולה לכפתור (כאן מתחיל הקוד שלך)
$button.Add_Click({
    $InputUser = $textBoxUser.Text.Trim()
    $InputOperator = $textBoxOperator.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($InputUser) -or [string]::IsNullOrWhiteSpace($InputOperator)) {
        [System.Windows.Forms.MessageBox]::Show("נא למלא את שני השדות.", "שגיאה")
        return
    }

    try {
        # ננסה לפי שם התחברות
        try {
            $ADUser = Get-ADUser -Identity $InputUser -Properties MemberOf, PrimaryGroupID, DisplayName -ErrorAction Stop
        } catch {
            $ADUser = Get-ADUser -Filter { Name -eq $InputUser } -Properties MemberOf, PrimaryGroupID -ErrorAction Stop
        }

        $Groups = Get-ADPrincipalGroupMembership -Identity $ADUser.SamAccountName | Select-Object Name, DistinguishedName, SID

        if (-not $Groups) {
            [System.Windows.Forms.MessageBox]::Show("לא נמצאו קבוצות עבור המשתמש '$($ADUser.SamAccountName)'.", "שגיאה")
            return
        }

        # יצירת תיקיית יעד אם לא קיימת
		$ADDisplayName = $ADUser.DisplayName
        $TargetPath = "\\evo-veeam22\D$\PST_Archive-NEW\deleted\$ADDisplayName"
        if (-not (Test-Path $TargetPath)) {
            New-Item -Path $TargetPath -ItemType Directory | Out-Null
        }

        # שמירת קובץ CSV עם הקבוצות
        $CsvPath = "$TargetPath\$($ADUser.SamAccountName).csv"
        $Groups | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

        # בדיקה שהקובץ נוצר וכולל יותר משורה אחת
        Start-Sleep -Seconds 1  # זמן קצר להבטחת כתיבה מלאה
        if ((Test-Path $CsvPath) -and ((Get-Content $CsvPath).Count -gt 1)) {

            # קבלת קבוצת Domain Users עם ה-PrimaryGroupToken
            $DomainUsersGroup = Get-ADGroup -Filter { Name -eq "Domain Users" } -Properties PrimaryGroupToken

            # שינוי הקבוצה הראשית ל-Domain Users אם צריך
            if ($ADUser.PrimaryGroupID -ne $DomainUsersGroup.PrimaryGroupToken) {
                Set-ADUser -Identity $ADUser.SamAccountName -Replace @{PrimaryGroupID = $DomainUsersGroup.PrimaryGroupToken}
            }

            # הסרת המשתמש מכל שאר הקבוצות מלבד Domain Users
            foreach ($Group in $Groups) {
                if ($Group.Name -ne "Domain Users") {
                    try {
                        Remove-ADGroupMember -Identity $Group.DistinguishedName -Members $ADUser.SamAccountName -Confirm:$false
                    } catch {
                        Write-Warning "שגיאה בהסרה מהקבוצה $($Group.Name): $_"
                    }
                }
            }
            Set-ADUser $ADUser -Description "Disabled by $InputOperator $((Get-Date).ToString('dd/MM/yy'))"
            Disable-ADAccount -Identity $ADUser
            $TargetOU = "OU=Evogene Disabled Users,OU=People,DC=Evo,DC=corp"
            Move-ADObject -Identity $ADUser.DistinguishedName -TargetPath $TargetOU
            [System.Windows.Forms.MessageBox]::Show("בוצע בהצלחה! הקובץ נשמר ב:`n$CsvPath`nוהקבוצות נמחקו למעט 'Domain Users'.", "הצלחה")
        } else {
            [System.Windows.Forms.MessageBox]::Show("הייצוא נכשל – הקובץ לא נוצר או ריק. המחיקה בוטלה.", "שגיאה")
        }

    } catch {
        [System.Windows.Forms.MessageBox]::Show("שגיאה: $_", "שגיאה")
    }
})

$form.Controls.Add($button)
$form.ShowDialog()
