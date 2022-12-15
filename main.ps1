# strings "C:\Windows\System32\ntoskrnl.exe" | findstr /i GlobalTimerResolutionRequests

# $dumpbin = "C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\VC\Tools\MSVC\14.29.30133\bin\Hostx86\x64\dumpbin.exe"
$dumpbin = ".\dumpbin.exe"
$depth = "3" # Set maximum search depth
$stringLength = 5 # Minimum string length (default is 3)

# Turn off Sleep Mode
Start-Process -FilePath powercfg -ArgumentList "-change -monitor-timeout-ac 0" -WindowStyle Hidden
Start-Process -FilePath powercfg -ArgumentList "-change -standby-timeout-ac 0" -WindowStyle Hidden

$ErrorActionPreference = 'SilentlyContinue'
if(-Not(Test-Path -Path "HKCU:\Software\Sysinternals")){
    New-Item -Path "HKCU:\Software" -Name "Sysinternals"
}
if(-Not(Test-Path -Path "HKCU:\Software\Sysinternals\Sysinternals")){
    New-Item -Path "HKCU:\Software\Sysinternals" -Name "Strings"
}
Set-ItemProperty -Path "HKCU:\Software\Sysinternals\Strings" -Name EulaAccepted -Value 1 -Type DWord -Force

$buildnum =@()
$buildnum += (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name CurrentMajorVersionNumber).CurrentMajorVersionNumber
$buildnum += (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name CurrentMinorVersionNumber).CurrentMinorVersionNumber
$buildnum += (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name CurrentBuild).CurrentBuild
$buildnum += (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name UBR).UBR #(Update Build Revision)

$tempfolder = ".\Temp\" + $($buildnum -join '_') + "\"

New-Item -Path $tempfolder -ItemType Directory -Force | Out-Null

$output = .\fd `
    "--absolute-path", `
    "--max-depth", $depth, `
    "--type", "f", `
    "--extension", "exe", `
    "--extension", "dll", `
    "--extension", "sys", `
    "--color", "never", `
    ".", `
    "C:\Windows"

$Array = $output.Split([Environment]::NewLine)
$all = $Array.Count
foreach ($target in $Array) {
    $cleanTarget = $target.Replace(":", "")
    $newPath = "$($tempfolder)$($cleanTarget)"
    $path = Split-Path -Path $newPath
    if (-not (Test-Path $path)) {
        New-Item -Path $path -ItemType Directory -Force | Out-Null
    }

    try {
        [array]$cmdOutput = .\strings.exe -nobanner -n $stringLength $target
        $cmdOutput = $cmdOutput | Sort-Object | Get-Unique
        Set-Content -LiteralPath "$($newPath).strings" -Value $cmdOutput
    }
    catch {
        write-warning $newPath
    }

    try {
        Set-Content -Path "$($newPath).coff" -Value (&$dumpbin "/ARCHIVEMEMBERS", "/CLRHEADER", "/DEPENDENTS", "/EXPORTS", "/IMPORTS", "/SUMMARY", "/SYMBOLS", $target)
    }
    catch {
        write-warning $newPath
    }

    $all -= 1
    if ($all % 5 -eq 0) {
        Write-Host $all
    }
}

#https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.archive/compress-archive?view=powershell-5.1
$zipFile = "$($buildnum -join '_').zip"
Compress-Archive -Path $tempfolder -DestinationPath ".\$zipFile" -CompressionLevel "NoCompression"
