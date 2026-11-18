$ErrorActionPreference = 'Stop'
$project = $PSScriptRoot
$line = [Environment]::NewLine
$format = Get-Content -LiteralPath (Join-Path $project 'src/index.html') -Raw -Encoding utf8
$style = Get-Content -LiteralPath (Join-Path $project 'src/style.css') -Raw -Encoding utf8
$license = Get-Content -LiteralPath (Join-Path $project 'vendor/JSZip-MIT.txt') -Raw -Encoding utf8
$scripts = @(
    'vendor/jszip.min.js'
    'src/citations.js'
    'src/word.js'
    'src/app.js'
) | ForEach-Object { Get-Content -LiteralPath (Join-Path $project $_) -Raw -Encoding utf8 }

$scriptBlocks = ($scripts | ForEach-Object { "<script>$line$_$line</script>" }) -join $line
$page = $format.Replace('<!-- styles -->', "<style>$line$style$line</style>")
$page = $page.Replace('<!-- scripts -->', "<!--$line$license$line-->$line$scriptBlocks")
$output = Join-Path $project 'dist'
New-Item -ItemType Directory -Force -Path $output | Out-Null
[IO.File]::WriteAllText((Join-Path $output 'index.html'), $page, [Text.UTF8Encoding]::new($false))
