# read-mail.ps1 — Outlook COM で指定フォルダのメールを JSON で標準出力へ書き出す（完全ローカル）
# 常駐エージェント（agent.js）から呼ばれる。EntryID を含めるので呼び出し側で処理済み判定ができる。
param(
  [string]$Folder = "受信トレイ",
  [int]$Days = 7,
  [int]$MaxCount = 100
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Find-Folder($parent, $name) {
  foreach ($f in $parent.Folders) {
    if ($f.Name -ieq $name) { return $f }
    $found = Find-Folder $f $name
    if ($null -ne $found) { return $found }
  }
  return $null
}

try {
  $outlook = New-Object -ComObject Outlook.Application
  $ns = $outlook.GetNamespace("MAPI")
  $inbox = $ns.GetDefaultFolder(6) # 6 = olFolderInbox

  if ($Folder -eq "受信トレイ" -or $Folder -ieq "inbox") {
    $target = $inbox
  } else {
    $target = Find-Folder $inbox $Folder
    if ($null -eq $target) {
      foreach ($store in $ns.Folders) {
        $target = Find-Folder $store $Folder
        if ($null -ne $target) { break }
      }
    }
  }
  if ($null -eq $target) {
    Write-Output (@{ error = "folder_not_found"; folder = $Folder } | ConvertTo-Json -Compress)
    exit 0
  }

  $since = (Get-Date).AddDays(-$Days)
  $items = $target.Items
  $items.Sort("[ReceivedTime]", $true) # 新しい順

  $result = New-Object System.Collections.ArrayList
  $count = 0
  foreach ($item in $items) {
    if ($count -ge $MaxCount) { break }
    try {
      if ($item.Class -ne 43) { continue } # 43 = olMail
      if ($item.ReceivedTime -lt $since) { break } # 新しい順なので期間外が出たら打ち切り
      $body = "$($item.Body)"
      if ($body.Length -gt 4000) { $body = $body.Substring(0, 4000) + "（以下省略）" }
      [void]$result.Add(@{
        id       = "$($item.EntryID)"
        subject  = "$($item.Subject)"
        from     = "$($item.SenderName)"
        received = $item.ReceivedTime.ToString("yyyy-MM-ddTHH:mm:ss")
        body     = $body
      })
      $count++
    } catch { continue }
  }

  Write-Output (@{ folder = "$($target.Name)"; mails = $result } | ConvertTo-Json -Compress -Depth 4)
} catch {
  Write-Output (@{ error = "outlook_error"; message = "$($_.Exception.Message)" } | ConvertTo-Json -Compress)
}
