﻿<# 
MIT License

Copyright (c) 2021 SATO Yusuke

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>
<# 
TODO:
  [v] シャットダウン時刻を毎分00秒にする
  [v] 指定した時刻にshutdown.exeが実行されるのではなく、指定した時刻にシャットダウンがされるようにする
  [v] shutdown.exeが走ったあとでも取り消しができるように
  [v] 過去時刻 or 現在時刻＋タイムアウトより前の時間指定はエラーにする
  [x] UIをXAMLで作り直す -> TimePickerはWinUIがないと使えない＆XAMLのTimePickerにスピンボタンがないため不適
  [ ] 時刻は15分刻みで上下するようスピンボタンの動作を変更 -> キーボードからの時刻入力と併用できない？
  [v] シャットダウンしたらまずい環境（管理者権限のないCitrix上のVMなど）のため、シャットダウンでなく再起動(shutdown.exe /r /t 300)ができるように
  [ ] 設定を保存できるようにする。自己書き換えできる？
  [ ] シャットダウンまでのタイムアウト時間を設定可能にする
  [ ] 日付けまたぎの処理ができるようにする（TimeSpan.TicksPerMinuteを使う？）
  [ ] TimePickerで時と分が連動して増加/減少するようにする
  [ ] マウスホイールで時間の増減ができるように、またはプルダウンでもよい
  [ ] shutdown.exeが走ったあとにタスクがちゃんと消えるようにする
  [v] ライセンスとファイルを一体化
  [ ] タスクスケジューラのルートフォルダにすでに「Shutdown」というタスクがあった場合に以下のエラーが出る問題の対処
  --------------------------------
  Register-ScheduledTask : アクセスが拒否されました。
  発生場所 C:\home\satob\git\Shutdown\ShutdownRegister.ps1:55 文字:9
  +         Register-ScheduledTask -TaskPath $TaskPath -TaskName $TaskNam ...
  +         ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      + CategoryInfo          : PermissionDenied: (PS_ScheduledTask:Root/Microsoft/...S_ScheduledTask) [Register-ScheduledTask], CimException
      + FullyQualifiedErrorId : HRESULT 0x80070005,Register-ScheduledTask
  --------------------------------
  [v] すでに登録されていた場合はエラーを返すのではなく削除してから再登録にする
#>
$TimeoutPeriod = 300

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$TaskName = "Shutdown"
$TaskPath = "\Shutdown\"

function UnregisterTask() {
    $Task = (Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue)
    if ($Task -ne $null) {
        Unregister-ScheduledTask -TaskName $Task.TaskName -AsJob
        Start-Process -FilePath shutdown.exe -ArgumentList "/a" -Wait
        [Windows.Forms.MessageBox]::Show("シャットダウンを解除しました。", "シャットダウン解除済み")
    } else {
        [Windows.Forms.MessageBox]::Show("シャットダウンは登録されていません。", "シャットダウン未登録")
    }
}

function RegisterTask {
    Param(
        $DateTime,
        [bool] $Reboot
    )

    $TriggerTime = [DateTime]($DateTime.AddSeconds(-$TimeoutPeriod).ToShortTimeString())

    if ((Get-Date) -ge $TriggerTime) {
        $LatestTime = (Get-Date).AddSeconds($TimeoutPeriod + 60).ToShortTimeString()
        [Windows.Forms.MessageBox]::Show("シャットダウン設定可能時刻は" + $LatestTime + "以降です。", "時刻指定エラー")
        return
    }

    $Task = (Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue)
    if ($Task -ne $null) {
        Unregister-ScheduledTask -TaskName $Task.TaskName -AsJob
        Start-Process -FilePath shutdown.exe -ArgumentList "/a" -Wait
    }

    if ($Reboot) {
        $TaskAction = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument ("/r /t " + $TimeoutPeriod)
    } else {
        $TaskAction = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument ("/s /t " + $TimeoutPeriod)
    }
    $TaskTrigger = New-ScheduledTaskTrigger -Once -At $TriggerTime
    Register-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger
    [Windows.Forms.MessageBox]::Show("シャットダウン (" + $DateTime.ToShortTimeString() + ") を登録しました。", "シャットダウン登録済み")
}

function Roundup5Minutes() {
        Param([DateTime] $DateTime)

        $Quantize = 5
        $minute = $DateTime.Minute

        $addedMinutes = ($Quantize - ($minute + $Quantize) % $Quantize)
        return $DateTime.AddMinutes($addedMinutes + 5)
}


# フォームの作成
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(300,120)
$Form.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedSingle
$Form.Text = "シャットダウン設定君"


# 時刻フォームの設定
$TimePicker = New-Object System.Windows.Forms.DateTimePicker
$TimePicker.Location = New-Object System.Drawing.Point(10,10)
$TimePicker.Size = New-Object System.Drawing.Size(80,24)
$TimePicker.Format = [Windows.Forms.DateTimePickerFormat]::Custom
$TimePicker.CustomFormat = "HH:mm"
$TimePicker.ShowUpDown = $true

# 時刻の文字が小さいので大きめにする
$CurrentFont = $TimePicker.Font
$NewFont = New-Object System.Drawing.Font($CurrentFont.Name, 14, $CurrentFont.Style, $CurrentFont.Unit)
$TimePicker.Font = $NewFont

# 時刻は最低でも5分後
$TimePicker.Value = (Roundup5Minutes(Get-Date))

$Form.Controls.Add($TimePicker)


# 再起動/シャットダウン ラジオボタンの設定
$RebootGroupBox = New-Object System.Windows.Forms.GroupBox
$RebootGroupBox.Location = New-Object System.Drawing.Point(100,4)
$RebootGroupBox.size = New-Object System.Drawing.Size(170,36)

$RebootRadio = New-Object System.Windows.Forms.RadioButton
$RebootRadio.Location = New-Object System.Drawing.Point(10,8)
$RebootRadio.size = New-Object System.Drawing.Size(60,25)
$RebootRadio.Checked = $True
$RebootRadio.Text = "再起動"

$ShutdownRadio = New-Object System.Windows.Forms.RadioButton
$ShutdownRadio.Location = New-Object System.Drawing.Point(75,8)
$ShutdownRadio.size = New-Object System.Drawing.Size(90,25)
$ShutdownRadio.Text = "シャットダウン"

$RebootGroupBox.Controls.AddRange(@($RebootRadio,$ShutdownRadio))
$Form.Controls.Add($RebootGroupBox)


# 登録ボタンの設定
$RegisterButton = New-Object System.Windows.Forms.Button
$RegisterButton.Location = New-Object System.Drawing.Point(100,46)
$RegisterButton.Size = New-Object System.Drawing.Size(80,25)
$RegisterButton.Text = "登録する"
$Form.Controls.Add($RegisterButton)

function RegisterButton_Click(){
    RegisterTask -DateTime $TimePicker.Value -Reboot $RebootRadio.Checked
}
$RegisterButton.Add_Click({RegisterButton_Click})


# 解除ボタンの設定
$UnregisterButton = New-Object System.Windows.Forms.Button
$UnregisterButton.Location = New-Object System.Drawing.Point(190,46)
$UnregisterButton.Size = New-Object System.Drawing.Size(80,25)
$UnregisterButton.Text = "解除する"
$Form.Controls.Add($UnregisterButton)
function UnregisterButton_Click(){
    UnregisterTask
}
$UnregisterButton.Add_Click({UnregisterButton_Click})


$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()

