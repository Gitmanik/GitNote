Param(
    [string]$Loc
)

#https://github.com/lukegackle/PowerShell-Self-Elevate-Keeping-Current-Directory/blob/master/Self%20Elevate%20Keeping%20Directory.ps1
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator'))
{
    $Arguments =  @(
        '-NoProfile',
        '-ExecutionPolicy Bypass',
        '-File',
        "`"$($MyInvocation.MyCommand.Path)`"",
        "\`"$(Get-Location)\\`""
    )
    Start-Process -FilePath PowerShell.exe -Verb RunAs -ArgumentList $Arguments
    Exit
}

if($Loc.Length -gt 1){
Set-Location $($Loc.Substring(1,$Loc.Length-1)).Trim()
}
$GitNoteDLL = Join-Path (Get-Location) 'GitNote\bin\x64\Debug\Gitmanik.GitNote.dll'

if (!(Test-Path $GitNoteDLL))
{
    Write-Host "Cannot find $GitNoteDLL. Make sure to build the solution first."
    Pause
    return
}

$GitNoteAssembly = [Reflection.Assembly]::Loadfile($GitNoteDLL)
$GitNoteVersion = $GitNoteAssembly.GetName().Version
Write-Host "GitNote version: $GitNoteVersion"

$Reg_Path = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Gitmanik.GitNote.dll"
if (!(Test-Path $Reg_Path))
{
    Write-Host "Cannot find registry path for GitNote DLL association. Make sure to install GitNote from installer first."
    Pause
    return
}


Set-ItemProperty $Reg_Path -Name 'Path' -Type String -Value $GitNoteDLL

$Reg_Path2 = "Registry::HKEY_CLASSES_ROOT\CLSID\{7562BB0E-F7F7-4B01-A3BA-9FD8C5711F14}\InprocServer32\" + $GitNoteVersion
if (!(Test-Path $Reg_Path2))
{
    Write-Host "Cannot find registry path for GitNote COM association. Make sure to install GitNote from installer first."
    Pause
    return
}

Set-ItemProperty $Reg_Path2 -Name 'CodeBase' -Type String -Value $GitNoteDLL

Write-Host "OneNote will now use Debug DLL instead of installed one. To revert run the installer again."

Pause