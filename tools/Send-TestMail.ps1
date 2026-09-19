<#
.SYNOPSIS
  Sends a test message through MailGateway, optionally with attachments and a custom From.

.EXAMPLE
  .\Send-TestMail.ps1 -Url http://localhost:5001/api/mail/send -ApiKey $env:MG_KEY -To you@piruli.com

.EXAMPLE
  .\Send-TestMail.ps1 -Url https://apps-thanos.central-office.info/mailgateway_AI/api/mail/send -ApiKey "..." `
      -To you@piruli.com -From noreply@piruli.com -FromName "Piruli" `
      -Attach "C:\temp\receipt.pdf", "C:\temp\photo.jpg"

  Quote every path (spaces in file names would otherwise split into extra arguments).
#>
param(
    [Parameter(Mandatory)] [string] $Url,
    [Parameter(Mandatory)] [string] $ApiKey,
    [Parameter(Mandatory)] [string] $To,
    [string] $From = "",
    [string] $FromName = "",
    [string] $Subject = "MailGateway test $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
    [string[]] $Attach = @(),
    [string] $InlineImage = ""     # path to a PNG/JPG shown inline in the body
)

$ErrorActionPreference = "Stop"

$attachments = New-Object System.Collections.Generic.List[object]

foreach ($path in @($Attach)) {
    if ([string]::IsNullOrWhiteSpace($path)) { continue }
    if (-not (Test-Path -LiteralPath $path)) { throw "Attachment not found: $path" }

    $bytes = [IO.File]::ReadAllBytes($path)
    $attachments.Add(@{
        name          = [IO.Path]::GetFileName($path)
        contentBase64 = [Convert]::ToBase64String($bytes)
    })
    Write-Host ("  attach: {0}  ({1:N0} bytes)" -f [IO.Path]::GetFileName($path), $bytes.Length)
}

$html = "<p>Test from <b>MailGateway</b> at $(Get-Date).</p>"

if ($InlineImage) {
    if (-not (Test-Path -LiteralPath $InlineImage)) { throw "Inline image not found: $InlineImage" }
    $bytes = [IO.File]::ReadAllBytes($InlineImage)
    $attachments.Add(@{
        name          = [IO.Path]::GetFileName($InlineImage)
        contentBase64 = [Convert]::ToBase64String($bytes)
        contentId     = "inline1"
        isInline      = $true
    })
    $html += '<p><img src="cid:inline1" alt="inline"></p>'
    Write-Host ("  inline: {0}  ({1:N0} bytes)" -f [IO.Path]::GetFileName($InlineImage), $bytes.Length)
}

$body = [ordered]@{
    appName     = "Send-TestMail.ps1"
    to          = $To
    subject     = $Subject
    htmlBody    = $html
    textBody    = "Test from MailGateway."
    attachments = @($attachments.ToArray())
}
if ($From)     { $body.from = $From }
if ($FromName) { $body.fromName = $FromName }

$json = ConvertTo-Json -InputObject $body -Depth 6 -Compress
Write-Host ("POST {0}  ({1:N0} KB, {2} attachment(s))" -f $Url, [Math]::Round($json.Length / 1KB), $attachments.Count)

try {
    $resp = Invoke-RestMethod -Method Post -Uri $Url -ContentType "application/json; charset=utf-8" `
        -Headers @{ "X-API-Key" = $ApiKey } -Body ([Text.Encoding]::UTF8.GetBytes($json)) -TimeoutSec 300
    $resp | ConvertTo-Json
}
catch {
    Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.ErrorDetails.Message) { Write-Host $_.ErrorDetails.Message }
    exit 1
}