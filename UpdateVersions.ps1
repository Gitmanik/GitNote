$GitNoteDLL = Join-Path (Get-Location) "../../GitNote/bin/x64/Debug/Gitmanik.GitNote.dll"
$GitNoteVDProj = "../Setup.vdproj"

Write-Host "VDProj file: $GitNoteVDProj"
Write-Host "Reading version from: $GitNoteDLL"



$GitNoteAssembly = [Reflection.Assembly]::Loadfile($GitNoteDLL)
$GitNoteVersion = $GitNoteAssembly.GetName().Version

$ProductCode = (New-Guid).ToString().ToUpper()
$UpgradeCode = (New-Guid).ToString().ToUpper()


Write-Host "GitNote version: $GitNoteVersion"
Write-Host "New installer Product Code: $ProductCode"
Write-Host "New installer Upgrade Code: $UpgradeCode"

$VDProjContent = (Get-Content $GitNoteVDProj)

$VDProjContent = $VDProjContent -replace '"ARPCOMMENTS" = "8:([0-9]+.[0-9]+.[0-9]+.[0-9]+)"', ('"ARPCOMMENTS" = "8:' + $GitNoteVersion + '"')
$VDProjContent = $VDProjContent -replace '"ProductCode" = "8:{[A-F0-9-]+}"', ('"ProductCode" = "8:{' + $ProductCode + '}"')
$VDProjContent = $VDProjContent -replace '"UpgradeCode" = "8:{[A-F0-9-]+}"', ('"UpgradeCode" = "8:{' + $UpgradeCode + '}"')
Set-Content -Path $GitNoteVDProj -Value $VDProjContent