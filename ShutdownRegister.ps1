<# 
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
  [v] シャットダウンしたらまずい環境（管理者権限のないCitrix上のVMなど）のため、シャットダウンでなく再起動(shutdown.exe /r /t 300)ができるように
  [ ] 設定を保存できるようにする。自己書き換えできる？
  [ ] シャットダウンまでのタイムアウト時間を設定可能にする
  [ ] 日付けまたぎの処理ができるようにする（TimeSpan.TicksPerMinuteを使う？）
  [v] TimePickerで時と分が連動して増加/減少するようにする
  [v] マウスホイールで時間の増減ができるように、またはプルダウンでもよい
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
  [v] HungAppTimeoutを待つようStop-Computer/Restart-Computerでシャットダウンするようにする
  [v] すでに登録されていた場合はエラーを返すのではなく削除してから再登録にする
  [v] 実行時にコマンドプロンプトを隠す
  [v] ウィンドウのアイコンを変更する
#>
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

$TimeoutPeriod = 300
$LatencyToForceShutdown = 20

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$ShutdownExeTaskName = "Shutdown"
$CmdletTaskName = "ShutdownCmdlet"
$TaskPath = "\Shutdown\"

function UnregisterTask() {
    $Task = (Get-ScheduledTask -TaskName $ShutdownExeTaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue)
    if ($Task -ne $null) {
        Unregister-ScheduledTask -TaskName $Task.TaskName -AsJob
        $CmdletTask = (Get-ScheduledTask -TaskName $CmdletTaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue)
        if ($CmdletTask -ne $null) {
            Unregister-ScheduledTask -TaskName $CmdletTask.TaskName -AsJob
        }
        Start-Process -FilePath shutdown.exe -ArgumentList "/a" -Wait
        [Windows.Forms.MessageBox]::Show("シャットダウンを解除しました。", "シャットダウン解除済み")
    } else {
        [Windows.Forms.MessageBox]::Show("シャットダウンは登録されていません。", "シャットダウン未登録")
    }
}

function RegisterTask {
    Param(
        $DateTime,
        [bool] $Reboot,
        [bool] $SlideToShutdown
    )

    # shutdown.exeの実行を開始する（全画面メッセージを表示する）時刻
    $ShutdownExeTriggerTime = [DateTime]($DateTime.AddSeconds(-$TimeoutPeriod).ToShortTimeString())

    # Stop-Computer/Restart-Computerを実行する時刻
    # 指定した時刻（分）ぴったりに実行開始する
    $ActualShutdownTriggerTime = [DateTime]($DateTime.ToShortTimeString())

    # 直近の時刻すぎる場合はエラーとする
    if ((Get-Date) -ge $ShutdownExeTriggerTime) {
        $LatestTime = (Get-Date).AddSeconds($TimeoutPeriod + 60).ToShortTimeString()
        [Windows.Forms.MessageBox]::Show("シャットダウン設定可能時刻は" + $LatestTime + "以降です。", "時刻指定エラー")
        return
    }

    # すでにシャットダウン登録済みの場合は登録解除する
    $ShutdownExeTask = (Get-ScheduledTask -TaskName $ShutdownExeTaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue)
    if ($ShutdownExeTask -ne $null) {
        Unregister-ScheduledTask -TaskName $ShutdownExeTask.TaskName -AsJob

        $CmdletTask = (Get-ScheduledTask -TaskName $CmdletTaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue)
        if ($CmdletTask -ne $null) {
            Unregister-ScheduledTask -TaskName $CmdletTask.TaskName -AsJob
        }
        # shutdown.exeが始まっている場合は停止する
        Start-Process -FilePath shutdown.exe -ArgumentList "/a" -Wait
    }

    # Stop-Computer/Restart-Computerを実行してからshutdown.exeによるシャットダウンが始まるまで余裕を持たせる
    if ($Reboot) {
        $ShutdownExeTaskAction = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument ("/r /t " + ($TimeoutPeriod + $LatencyToForceShutdown))
        $CmdletTaskAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "Restart-Computer"
    } elseif ($SlideToShutdown) {
        $ShutdownExeTaskAction = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument ("/s /t " + ($TimeoutPeriod + $LatencyToForceShutdown))
        $CmdletTaskAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "shutdown.exe /a; SlideToShutdown.exe"
    } else {
        $ShutdownExeTaskAction = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument ("/s /t " + ($TimeoutPeriod + $LatencyToForceShutdown))
        $CmdletTaskAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "Stop-Computer"
    }
    $ShutdownExeTaskTrigger = New-ScheduledTaskTrigger -Once -At $ShutdownExeTriggerTime
    $CmdletTaskTrigger = New-ScheduledTaskTrigger -Once -At $ActualShutdownTriggerTime

    $ScheduledTaskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries
    Register-ScheduledTask -TaskPath $TaskPath -TaskName $ShutdownExeTaskName -Action $ShutdownExeTaskAction -Trigger $ShutdownExeTaskTrigger -Settings $ScheduledTaskSettings
    Register-ScheduledTask -TaskPath $TaskPath -TaskName $CmdletTaskName -Action $CmdletTaskAction -Trigger $CmdletTaskTrigger -Settings $ScheduledTaskSettings
    [Windows.Forms.MessageBox]::Show("シャットダウン (" + $DateTime.ToShortTimeString() + ") を登録しました。", "シャットダウン登録済み")
}

function Roundup5Minutes() {
        Param([DateTime] $DateTime)

        $Quantize = 5
        $minute = $DateTime.Minute

        $addedMinutes = ($Quantize - ($minute + $Quantize) % $Quantize)
        return $DateTime.AddMinutes($addedMinutes + 5)
}


# コマンドプロンプトを隠す
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

# フォームの作成
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(480,130)
$Form.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedSingle
$Form.Text = "シャットダウン設定君"

# How to make gzip compressed Base64 string
# $byteArray = (Get-Content -Encoding Byte -Raw .\ShutdownRegister.ico)
# [System.IO.MemoryStream] $output = New-Object System.IO.MemoryStream
# $gzipStream = New-Object System.IO.Compression.GzipStream $output, ([IO.Compression.CompressionMode]::Compress)
# $gzipStream.Write( $byteArray, 0, $byteArray.Length )
# $gzipStream.Close()
# $output.Close()
# [System.Convert]::ToBase64String($output.ToArray()) | Set-Clipboard

$FormIconGZipBase64 = 'H4sIAAAAAAAEAJ1VT4jjZBx9ww7M4r+Zk6dgC4oICg6echBTBEGQZYel0GPn5E3akwcRUxSsh0JGBK9fD6v2ljnZY7JMYdGDGfAioiQLAz3MITl4qG7t86X/Jp1md8AvfUn4ve/35f3e90sK7OioVqHzbQxuAy8CeE1QCDUs4k8fnfkPnc7ikh+6CcMQ4R8hsp+/V/QZ4VnhOeF54QVhv/CUY8EV+kIoJMIMlkXYNlGvE+020esRgwExGhFJwpvSYdGCTRt11tFmGz32MOAAI46QMLkpHbQs0LbBeh1st8FeDxwMwNEITDQrN66mFY61gqsV+v28cC0gbib9lH6hLrSFnjAQRkIi3JD+FWm9S9ohpZ/ST+mn9FP6Oddflf6a9B9Lvyv9fekPpT+R/pn0U/oF+Uf5R/lH+Uf5R/l314vQNDFafgo3ILyfCPMr4f9JBGP5W+ug0vDgtAyaXR+uCWCGEYIoRjxOQe0zvQpoHNBvgoELRgaMAzCN0VFtXqMC03Lgd5sIjItoaBBHAdJxDKpXWPFARzlNH3SVZyKtIy5OxdbgVRowTgt+s4vANYjMEHEQIY3HyAVUKH2UPkofpY/SR+ljilxAxWvAMS00/a5qNDDREEEcIU7HeP1D4u2PiTtfEs1vtUc/EJ8PiW8eEt/9JnWq31P9RvX7qj9Q/ZHqj1V/qvo7ssij6qPqo+qj6pOWWDpSaVA6GvKnJX+68sfIn6H8ieTPWP7gnprsg1eA994HnI8A+2vgrR+BN34HXtXiR8Inwn3i4BeiNiFa6pu+cC4cHByoh6o4PDxUH9VwdHSkXjrWu9JWP7k4OTlRT/Vxeno6fx/Pz8/VWwmyLFN/5R0+Ef4VpsJj4R/hb2SPifvK6+QvsnYp/6nhFpf8yEP5qWyM87EdthZjvBqlbE5f5mNzgnVFX65GYUIZXZiwJOfSSvg1u6YvihPW7Jq/uLi4xmOTvyzw1hav7MUCax6b/EpBzhfTFxOWDy9JX/KXT+ax5d81HgX/rTIe6/2zyvnljFW8hEcxbpVPuAoX+eLd6t7avN0K4kn8Ux61KWrdH4XLhuZVf61P10va5kvLveJLDVnz2+x8xpIvZ69X/X94bNHkxGGGfYbY09fglr4oO9M97ExeArLPBAUS4ZHwgNg5m6Nz64yhkAgTgWcPyC8ezdxuMnX2sqmzn336zv7EsivT3aoz26263K3qL+9l8q83yamzwMzlf5Cg6LO+CAAA'
$FormIconGZip = [System.Convert]::FromBase64String($FormIconGZipBase64)
$FormIconGZipMemoryStream = New-Object System.IO.MemoryStream(, $FormIconGZip)
$FormIconMemoryStream = New-Object System.IO.MemoryStream
$GZipStream = New-Object System.IO.Compression.GzipStream $FormIconGZipMemoryStream, ([IO.Compression.CompressionMode]::Decompress)
$GZipStream.CopyTo( $FormIconMemoryStream )
$GZipStream.Close()
$FormIconGZipMemoryStream.Close()
$Form.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($FormIconMemoryStream).GetHIcon()))

# 時刻フォームの設定
$TimePicker = New-Object System.Windows.Forms.DateTimePicker
$TimePicker.Location = New-Object System.Drawing.Point(12,12)
$TimePicker.Size = New-Object System.Drawing.Size(95,29)
$TimePicker.Format = [Windows.Forms.DateTimePickerFormat]::Custom
$TimePicker.CustomFormat = "HH:mm"
$TimePicker.ShowUpDown = $true
$TimePickerWheelHandler = {
    if ($_.Delta -gt 0) {
        [System.Windows.Forms.SendKeys]::Send('{UP}')
    } else {
        [System.Windows.Forms.SendKeys]::Send('{DOWN}')
    }
}
$TimePicker.add_MouseWheel($TimePickerWheelHandler)

# 時刻の文字が小さいので大きめにする
$CurrentFont = $TimePicker.Font
$NewFont = New-Object System.Drawing.Font($CurrentFont.Name, 20, $CurrentFont.Style, $CurrentFont.Unit)
$TimePicker.Font = $NewFont

# 時刻は最低でも5分後
$TimePicker.Value = (Roundup5Minutes(Get-Date))
$Form.Controls.Add($TimePicker)

# 値の変更時に分と時を連携する
$Script:CurrentTimePickerValue = $TimePicker.Value
$Script:NowChangingHours = $False
$TimePickerValueChangeHandler = {
    Param($Sender)
    if ($Script:CurrentTimePickerValue.Minute -eq 59 -and $Sender.Value.Minute -eq 0 -and $Script:NowChangingHours -eq $False) {
        $Script:NowChangingHours = $True
        $Sender.Value = $Sender.Value.AddHours(1)
    } elseif ($Script:CurrentTimePickerValue.Minute -eq 0 -and $Sender.Value.Minute -eq 59 -and $Script:NowChangingHours -eq $False) {
        $Script:NowChangingHours = $True
        $Sender.Value = $Sender.Value.AddHours(-1)
    }
    $Script:NowChangingHours = $False
    $Script:CurrentTimePickerValue = $Sender.Value
}
$TimePicker.add_ValueChanged($TimePickerValueChangeHandler)

# 再起動/シャットダウン ラジオボタンの設定
$RebootGroupBox = New-Object System.Windows.Forms.GroupBox
$RebootGroupBox.Location = New-Object System.Drawing.Point(118,9)
$RebootGroupBox.size = New-Object System.Drawing.Size(330,36)

# ツールチップの作成
$ToolTip = New-Object System.Windows.Forms.ToolTip
$ToolTip.AutoPopDelay = 10000
$ToolTip.InitialDelay = 500
$ToolTip.ReshowDelay = 100
$ToolTip.ShowAlways = $true

$RebootRadio = New-Object System.Windows.Forms.RadioButton
$RebootRadio.Location = New-Object System.Drawing.Point(10,8)
$RebootRadio.size = New-Object System.Drawing.Size(90,25)
$RebootRadio.Checked = $True
$RebootRadio.Text = "自動再起動"

# 自動再起動ラジオボタンにツールチップを設定
$ToolTip.SetToolTip($RebootRadio, "指定した時刻に自動的に再起動を開始します。`n再起動開始後のキャンセルはできません。`nキャンセルする場合は登録解除を実行してください。")

$ShutdownRadio = New-Object System.Windows.Forms.RadioButton
$ShutdownRadio.Location = New-Object System.Drawing.Point(100,8)
$ShutdownRadio.size = New-Object System.Drawing.Size(110,25)
$ShutdownRadio.Text = "自動シャットダウン"

# 自動シャットダウンラジオボタンにツールチップを設定
$ToolTip.SetToolTip($ShutdownRadio, "指定した時刻に自動的にシャットダウンを開始します。`nシャットダウン開始後のキャンセルはできません。`nキャンセルする場合は登録解除を実行してください。")

$SlideToShutdownRadio = New-Object System.Windows.Forms.RadioButton
$SlideToShutdownRadio.Location = New-Object System.Drawing.Point(215,8)
$SlideToShutdownRadio.Size = New-Object System.Drawing.Size(110,25)
$SlideToShutdownRadio.Text = "手動シャットダウン"

# 手動シャットダウンラジオボタンにツールチップを設定
$ToolTip.SetToolTip($SlideToShutdownRadio, "指定した時刻にシャットダウンを開始します。`nシャットダウン開始後でも、画面を上にスライドすることでにキャンセルが可能です。")

$RebootGroupBox.Controls.AddRange(@($RebootRadio,$ShutdownRadio,$SlideToShutdownRadio))
$Form.Controls.Add($RebootGroupBox)


# 登録ボタンの設定
$RegisterButton = New-Object System.Windows.Forms.Button
$RegisterButton.Location = New-Object System.Drawing.Point(130,56)
$RegisterButton.Size = New-Object System.Drawing.Size(80,25)
$RegisterButton.Text = "登録する"
$Form.Controls.Add($RegisterButton)

function RegisterButton_Click(){
    RegisterTask -DateTime $TimePicker.Value -Reboot $RebootRadio.Checked -SlideToShutdown $SlideToShutdownRadio.Checked
}
$RegisterButton.Add_Click({RegisterButton_Click})


# 解除ボタンの設定
$UnregisterButton = New-Object System.Windows.Forms.Button
$UnregisterButton.Location = New-Object System.Drawing.Point(220,56)
$UnregisterButton.Size = New-Object System.Drawing.Size(80,25)
$UnregisterButton.Text = "解除する"
$Form.Controls.Add($UnregisterButton)
function UnregisterButton_Click(){
    UnregisterTask
}
$UnregisterButton.Add_Click({UnregisterButton_Click})

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()

