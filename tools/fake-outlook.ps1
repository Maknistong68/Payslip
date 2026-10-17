# Test harness: runs a Windows "Send all" pack against a stand-in for Outlook and writes calls.json.
param([string]$Pack, [string]$Mode = 'send', [string]$FailFor = '')
$global:FakeCalls = New-Object System.Collections.ArrayList
$global:FakeFailFor = $FailFor
$account = [pscustomobject]@{ SmtpAddress = 'payroll@example.com' }
$ns = [pscustomobject]@{ Accounts = @($account) }
$ns | Add-Member ScriptMethod SendAndReceive { param($x) [void]$global:FakeCalls.Add(@{ op = 'sendreceive' }) }
$outlook = [pscustomobject]@{ NS = $ns }
$outlook | Add-Member ScriptMethod GetNamespace { param($n) return $this.NS }
$outlook | Add-Member ScriptMethod CreateItem {
  param($t)
  $m = [pscustomobject]@{ To = ''; Subject = ''; Body = ''; SendUsingAccount = $null; Files = (New-Object System.Collections.ArrayList) }
  $att = [pscustomobject]@{ Owner = $m }
  $att | Add-Member ScriptMethod Add { param($p) if (-not (Test-Path $p)) { throw "no file $p" }; [void]$this.Owner.Files.Add([string]$p) }
  $m | Add-Member NoteProperty Attachments $att
  $m | Add-Member ScriptMethod Send {
    if ($global:FakeFailFor -and $this.To -eq $global:FakeFailFor) { throw 'mailbox refused' }
    [void]$global:FakeCalls.Add(@{ op = 'send'; to = $this.To; subject = $this.Subject; body = $this.Body; files = @($this.Files) })
  }
  $m | Add-Member ScriptMethod Save { [void]$global:FakeCalls.Add(@{ op = 'save'; to = $this.To; files = @($this.Files) }) }
  return $m
}
$script = Join-Path $Pack 'send-all.ps1'
switch ($Mode) {
  'check'  { & $script -Check -Yes }
  'test'   { & $script -Test -Yes -Outlook $outlook }
  'drafts' { & $script -Drafts -Yes -Delay 0 -Outlook $outlook }
  default  { & $script -Yes -Delay 0 -Outlook $outlook }
}
ConvertTo-Json -InputObject @($global:FakeCalls) -Depth 5 | Set-Content -Path (Join-Path $Pack 'calls.json') -Encoding UTF8
