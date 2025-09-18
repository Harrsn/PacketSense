Get-ChildItem 'C:\ProgramData\SecureNet\NetMon\logs\*.csv' |
  Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
  Remove-Item -Force
