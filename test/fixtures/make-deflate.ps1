param([string]$arg)
$ErrorActionPreference = 'Stop'
$pair = $arg.Split('|')
$src = $pair[0]
$dest = $pair[1]
$memo = Join-Path $src ([char]0x65E5 + [char]0x672C + [char]0x8A9E + [char]0x30E1 + [char]0x30E2 + '.txt')
[System.IO.File]::WriteAllText($memo, 'jp')
Add-Type -AssemblyName System.IO.Compression.FileSystem
if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Force }
[System.IO.Compression.ZipFile]::CreateFromDirectory($src, $dest)
