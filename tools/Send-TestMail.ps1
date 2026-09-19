<#
.SYNOPSIS
  Sends a test message through MailGateway, optionally with attachments and a custom From.

.EXAMPLE
  .\Send-TestMail.ps1 -Url http://localhost:5001/api/mail/send -ApiKey $env:MG_KEY -To you@piruli.com

.EXAMPLE
  .\Send-TestMail.ps1 -Url https://apps.central-office.info/mailgateway/api/mail/send -ApiKey ... `
      -To you@aps-spr.org -From noreply@pasmeeting.org -FromName "PAS Meeting" `
      -Attach C:\temp\receipt.pdf, C:\temp\photo.jpg
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

$attachments = @()
foreach ($path in $Attach) {
    $attachments += @{
        name          = [IO.Path]::GetFileName($path)
        contentBase64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($path))
    }
}

$html = "<p>Test from <b>MailGateway</b> at $(Get-Date).</p>"
if ($InlineImage) {
    $attachments += @{
        name          = [IO.Path]::GetFileName($InlineImage)
        contentBase64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($InlineImage))
        contentId     = "inline1"
        isInline      = $true
    }
    $html += '<p><img src="cid:inline1" alt="inline"></p>'
}

$body = @{
    appName     = "Send-TestMail.ps1"
    to          = $To
    subject     = $Subject
    htmlBody    = $html
    textBody    = "Test from MailGateway."
    attachments = $attachments
}
if ($From)     { $body.from = $From }
if ($FromName) { $body.fromName = $FromName }

$json = $body | ConvertTo-Json -Depth 5
Write-Host "POST $Url  ($([Math]::Round($json.Length / 1KB)) KB)"

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
