# setup.ps1 — Win XPS 打印节点初始化（autounattend FirstLogonCommands 调用）
$ErrorActionPreference = 'Continue'
$dir = 'C:\printnode'
New-Item -ItemType Directory -Force -Path $dir, "$dir\jobs", "$dir\print" | Out-Null
Copy-Item "$PSScriptRoot\PrintParseService.cs" "$dir\" -Force
"start $(Get-Date)" | Out-File "$dir\setup-ps.log"

# --- RDP ---
Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -Value 0
netsh advfirewall firewall add rule name="RDP-3389" dir=in action=allow protocol=TCP localport=3389 | Out-Null

# --- 服务端口 ---
netsh advfirewall firewall add rule name="PrintNode-8090" dir=in action=allow protocol=TCP localport=8090 | Out-Null
netsh http add urlacl url=http://+:8090/ user=Everyone | Out-Null

# --- 电源：高性能，永不睡眠 ---
powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
powercfg /change standby-timeout-ac 0
powercfg /change monitor-timeout-ac 0

# --- 关 Windows Update 自动更新/驱动自动下载 ---
New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' -Force | Out-Null
Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' -Name NoAutoUpdate -Value 1
Set-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DriverSearching' -Name SearchOrderConfig -Value 0

# --- XPS 打印机固定到文件端口（Local Port 技巧：打印直接写文件不弹窗） ---
$port = 'C:\printnode\print\output.xps'
if (-not (Get-PrinterPort -Name $port -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name $port
}
Set-Printer -Name 'Microsoft XPS Document Writer' -PortName $port

# --- 编译服务（零外部依赖，用系统自带 csc.exe） ---
$csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $csc /nologo /target:exe /out:"$dir\PrintParseService.exe" "$dir\PrintParseService.cs" `
  /lib:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\WPF" `
  /r:System.dll /r:System.Core.dll /r:System.Web.Extensions.dll `
  /r:WindowsBase.dll /r:PresentationCore.dll /r:PresentationFramework.dll `
  /r:ReachFramework.dll /r:System.Xaml.dll 2>&1 | Out-File -Append "$dir\setup-ps.log"

if (Test-Path "$dir\PrintParseService.exe") {
    # --- 开机自启（deploy AutoLogon 时以最高权限运行） ---
    schtasks /create /f /tn PrintParseService /tr '"C:\printnode\PrintParseService.exe"' /sc onlogon /ru deploy /rl HIGHEST | Out-Null
    schtasks /run /tn PrintParseService | Out-Null
    "compile+task OK $(Get-Date)" | Out-File -Append "$dir\setup-ps.log"
} else {
    "COMPILE FAILED $(Get-Date)" | Out-File -Append "$dir\setup-ps.log"
}

"done $(Get-Date)" | Out-File "$dir\SETUP_DONE.txt"
