
$version = "v" + (Get-Date -Format "yyyyMMdd") 
& upx .\target\release\give_back_to_ceasar_or_god.exe
&7z a  "target\release\Schrödinger's Office_$version.zip"  assets/*.exe test/*
&7z u -spf2  "target\release\Schrödinger's Office_$version.zip" .\assets\register.cmd .\README.md .\target\release\give_back_to_ceasar_or_god.exe
