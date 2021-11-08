[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$TaskName = "Shutdown"
$TaskPath = "\Shutdown\"

function UnregisterTask() {
    $Task = (Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue)
    if ($Task -ne $null) {
        Unregister-ScheduledTask -TaskName $Task.TaskName -AsJob
        [Windows.Forms.MessageBox]::Show("シャットダウンを解除しました。", "シャットダウン解除済み")
    } else {
        [Windows.Forms.MessageBox]::Show("シャットダウンは登録されていません。", "シャットダウン未登録")
    }
}

function RegisterTask {
    Param($DateTime)

    $Task = (Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue)
    if ($Task -eq $null) {
        $TaskAction = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument "/s /t 300"
        $TaskTrigger = New-ScheduledTaskTrigger -Once -At $DateTime
        Register-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger
        [Windows.Forms.MessageBox]::Show("シャットダウンを登録しました。" + $DateTime.ToShortTimeString(), "シャットダウン登録済み")
    } else {
        [Windows.Forms.MessageBox]::Show("すでにシャットダウンが予定されています。", "シャットダウン登録済み")
    }
}


# フォームの作成
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(300,84)
$Form.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedSingle
$Form.Text = "シャットダウン設定君"


# 時刻フォームの設定
$TimePicker = New-Object System.Windows.Forms.DateTimePicker
$TimePicker.Location = New-Object System.Drawing.Point(10,10)
$TimePicker.Size = New-Object System.Drawing.Size(80,24)
$TimePicker.Format = [Windows.Forms.DateTimePickerFormat]::Custom
# 時刻の文字が小さいので大きめにする
$CurrentFont = $TimePicker.Font
$NewFont = New-Object System.Drawing.Font($CurrentFont.Name, 14, $CurrentFont.Style, $CurrentFont.Unit)
$TimePicker.Font = $NewFont
$TimePicker.CustomFormat = "HH:mm"
$TimePicker.ShowUpDown = $true
$Form.Controls.Add($TimePicker)


# 登録ボタンの設定
$RegisterButton = New-Object System.Windows.Forms.Button
$RegisterButton.Location = New-Object System.Drawing.Point(100,10)
$RegisterButton.Size = New-Object System.Drawing.Size(80,25)
$RegisterButton.Text = "登録する"
$Form.Controls.Add($RegisterButton)

function RegisterButton_Click(){
    RegisterTask($TimePicker.Value)
}
$RegisterButton.Add_Click({RegisterButton_Click})


# 解除ボタンの設定
$UnregisterButton = New-Object System.Windows.Forms.Button
$UnregisterButton.Location = New-Object System.Drawing.Point(190,10)
$UnregisterButton.Size = New-Object System.Drawing.Size(80,25)
$UnregisterButton.Text = "解除する"
$Form.Controls.Add($UnregisterButton)
function UnregisterButton_Click(){
    UnregisterTask
}
$UnregisterButton.Add_Click({UnregisterButton_Click})


$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
