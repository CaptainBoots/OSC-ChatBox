param(
    [string]$SvgPath = "Boot's-ToolBox.svg",
    [string]$IcoPath = "Boot's-ToolBox-256.ico",
    [switch]$KeepTemp
)

$ErrorActionPreference = "Stop"

function Resolve-FullPath {
    param(
        [Parameter(Mandatory = $true)][string]$PathValue,
        [Parameter(Mandatory = $true)][string]$BaseDir
    )

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return [System.IO.Path]::GetFullPath($PathValue)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $BaseDir $PathValue))
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$svg = Resolve-FullPath -PathValue $SvgPath -BaseDir $scriptDir
$ico = Resolve-FullPath -PathValue $IcoPath -BaseDir $scriptDir

if (-not (Test-Path -LiteralPath $svg)) {
    throw "SVG file not found: $svg"
}

$magick = (Get-Command magick -ErrorAction Stop).Source
$size = 256
$innerSize = 228
$iconSizes = "16,24,32,48,64,128,256"

$tempRoot = Join-Path $scriptDir ".icon-build"
$tempDir = Join-Path $tempRoot ([Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

try {
    $png = Join-Path $tempDir "icon-master-256.png"
    $density = 3072

    & $magick `
        -background none `
        -define svg:background-color=none `
        -density $density `
        $svg `
        -alpha set `
        -trim +repage `
        -filter LanczosSharp `
        -resize ("{0}x{0}" -f $innerSize) `
        -gravity center `
        -background none `
        -extent ("{0}x{0}" -f $size) `
        -colorspace sRGB `
        -depth 8 `
        -strip `
        $png

    $icoDir = Split-Path -Parent $ico
    if ($icoDir -and -not (Test-Path -LiteralPath $icoDir)) {
        New-Item -ItemType Directory -Path $icoDir -Force | Out-Null
    }

    # Build a multi-size ICO from the centered 256 master so Windows picks
    # crisp native layers at each desktop scale.
    & $magick `
        $png `
        -define ("icon:auto-resize={0}" -f $iconSizes) `
        $ico

    Write-Host "Created icon: $ico"
    & $magick identify -format "%f[%p] %wx%h opaque=%[opaque] alpha=%[channels]`n" $ico
}
finally {
    if (-not $KeepTemp -and (Test-Path -LiteralPath $tempDir)) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }
}
