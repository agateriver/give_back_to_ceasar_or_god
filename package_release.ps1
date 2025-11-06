
$version = "v" + (Get-Date -Format "yyyyMMdd") 
$releasePath = "target\release\Schrödinger's Office_$version.zip"

# Check if the release package exists and delete it
if (Test-Path -Path $releasePath -PathType Leaf) {
    Remove-Item -Path $releasePath -Force
    Write-Host "Old '$releasePath' has been deleted." -ForegroundColor Yellow
}

& upx -qqq .\target\release\give_back_to_ceasar_or_god.exe
& 7z a -tzip -bso0 "target\release\Schrödinger's Office_$version.zip"  assets/*.exe test/*
& 7z u -spf2  -bso0 "target\release\Schrödinger's Office_$version.zip" .\assets\register.cmd .\README.md .\target\release\give_back_to_ceasar_or_god.exe
Write-Host "Package created: target\release\Schrödinger's Office_$version.zip"  -ForegroundColor Green
