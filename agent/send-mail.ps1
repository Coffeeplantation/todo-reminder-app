# send-mail.ps1 — Outlook COM でリマインダーメールを送信または下書き作成する（完全ローカル）
# 常駐エージェント（agent.js）から呼ばれる。入力はJSONファイル（{ send: bool, items: [{to, cc, subject, body, from}] }）。
# send=true なら実際に送信（.Send）、false なら Outlook の下書きフォルダに保存（.Save）。
param(
  [Parameter(Mandatory = $true)][string]$JsonPath
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {
  $payload = Get-Content -Raw -Encoding UTF8 $JsonPath | ConvertFrom-Json
  $outlook = New-Object -ComObject Outlook.Application

  $sent = 0
  $errors = New-Object System.Collections.ArrayList
  foreach ($m in $payload.items) {
    try {
      $mail = $outlook.CreateItem(0) # 0 = olMailItem
      $mail.To = "$($m.to)"
      if ($m.cc) { $mail.CC = "$($m.cc)" }
      $mail.Subject = "$($m.subject)"
      $mail.Body = "$($m.body)"
      if ($m.from) { $mail.SentOnBehalfOfName = "$($m.from)" } # 複数アカウント時の差出人指定
      if ($payload.send) { $mail.Send() } else { $mail.Save() } # Save = 下書きフォルダへ
      $sent++
    } catch {
      [void]$errors.Add("$($m.subject): $($_.Exception.Message)")
    }
  }
  Write-Output (@{ ok = $sent; mode = $(if ($payload.send) { "send" } else { "draft" }); errors = $errors } | ConvertTo-Json -Compress)
} catch {
  Write-Output (@{ error = "outlook_error"; message = "$($_.Exception.Message)" } | ConvertTo-Json -Compress)
}
