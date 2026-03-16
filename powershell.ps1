
$ReportPath = "C:\Temp\SlowPC_Diagnostic_Report.txt"
New-Item -ItemType Directory -Force -Path "C:\Temp" | Out-Null

"===== SLOW PC DIAGNOSTIC REPORT =====" | Out-File $ReportPath
"Generated: $(Get-Date)" | Out-File $ReportPath -Append
"" | Out-File $ReportPath -Append

"===== SYSTEM INFORMATION =====" | Out-File $ReportPath -Append
Get-ComputerInfo | Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== CPU USAGE =====" | Out-File $ReportPath -Append
(Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples |
Select-Object CookedValue | Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== MEMORY STATUS =====" | Out-File $ReportPath -Append
$os = Get-CimInstance Win32_OperatingSystem
$totalMem = [math]::Round($os.TotalVisibleMemorySize/1MB,2)
$freeMem = [math]::Round($os.FreePhysicalMemory/1MB,2)
$usedMem = [math]::Round($totalMem - $freeMem,2)

"Total RAM (GB): $totalMem" | Out-File $ReportPath -Append
"Used RAM (GB): $usedMem" | Out-File $ReportPath -Append
"Free RAM (GB): $freeMem" | Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== DISK USAGE =====" | Out-File $ReportPath -Append
Get-PSDrive -PSProvider FileSystem | Select Name,
@{Name="UsedGB";Expression={[math]::Round($_.Used/1GB,2)}},
@{Name="FreeGB";Expression={[math]::Round($_.Free/1GB,2)}} |
Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== DISK HEALTH =====" | Out-File $ReportPath -Append
wmic diskdrive get model,status | Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== DISK QUEUE LENGTH =====" | Out-File $ReportPath -Append
(Get-Counter '\PhysicalDisk(_Total)\Avg. Disk Queue Length').CounterSamples |
Select CookedValue | Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== TOP CPU PROCESSES =====" | Out-File $ReportPath -Append
Get-Process | Sort CPU -Descending | Select -First 10 Name,CPU,Id |
Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== TOP MEMORY PROCESSES =====" | Out-File $ReportPath -Append
Get-Process | Sort WorkingSet -Descending |
Select -First 10 Name,
@{Name="MemoryMB";Expression={[math]::Round($_.WorkingSet/1MB,2)}},Id |
Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== TOP DISK PROCESSES =====" | Out-File $ReportPath -Append
Get-Process | Sort IOReadBytes -Descending |
Select -First 10 Name,IOReadBytes,IOWriteBytes |
Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== STARTUP PROGRAMS =====" | Out-File $ReportPath -Append
Get-CimInstance Win32_StartupCommand |
Select Name,Command,Location |
Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== WINDOWS UPDATE HISTORY =====" | Out-File $ReportPath -Append
Get-WinEvent -LogName System |
Where-Object {$_.ProviderName -like "*WindowsUpdate*"} |
Select -First 20 TimeCreated,Id,Message |
Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== RECENT SYSTEM ERRORS =====" | Out-File $ReportPath -Append
Get-EventLog -LogName System -EntryType Error -Newest 20 |
Select TimeGenerated,Source,EventID,Message |
Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== RECENT APPLICATION ERRORS =====" | Out-File $ReportPath -Append
Get-EventLog -LogName Application -EntryType Error -Newest 20 |
Select TimeGenerated,Source,EventID,Message |
Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== INSTALLED DRIVERS =====" | Out-File $ReportPath -Append
Get-WmiObject Win32_PnPSignedDriver |
Select DeviceName,DriverVersion,DriverProviderName |
Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== SERVICES WITH AUTO START =====" | Out-File $ReportPath -Append
Get-Service | Where {$_.StartType -eq "Automatic"} |
Select Name,Status |
Out-File $ReportPath -Append

"" | Out-File $ReportPath -Append
"===== REPORT COMPLETE =====" | Out-File $ReportPath -Append

Write-Host "Diagnostic report generated at $ReportPath"

