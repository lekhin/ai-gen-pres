<#
 Render all slides of a .pptx to PNG via PowerPoint COM.
 Usage: powershell -ExecutionPolicy Bypass -File pptx_render.ps1 -Src file.pptx -Dst outdir [-Width 1500]
 Requires Microsoft PowerPoint installed. Width <= 1500 px is recommended
 (vision model reads images better when they are pre-downscaled).
 Keep this file ASCII-only: PowerShell 5.1 misparses UTF-8 without BOM.
#>
param(
    [Parameter(Mandatory = $true)][string]$Src,
    [Parameter(Mandatory = $true)][string]$Dst,
    [int]$Width = 1500,
    [switch]$Svg
)

$ErrorActionPreference = 'Stop'
$Src = (Resolve-Path $Src).Path
if (-not (Test-Path $Dst)) { New-Item -ItemType Directory -Force -Path $Dst | Out-Null }
$Dst = (Resolve-Path $Dst).Path

$pres = $null
$ppt = New-Object -ComObject PowerPoint.Application
try {
    # ReadOnly=$true, Untitled=$false, WithWindow=$false
    $pres = $ppt.Presentations.Open($Src, $true, $false, $false)
    $h = [int]($Width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $out = Join-Path $Dst ("slide{0}.png" -f $i)
        $pres.Slides.Item($i).Export($out, 'PNG', $Width, $h) | Out-Null
        Write-Output ("rendered: {0}" -f $out)
        if ($Svg) {
            $svgFile = Join-Path $Dst ("slide{0}.svg" -f $i)
            try {
                $pres.Slides.Item($i).Export($svgFile, 'SVG', 0, 0) | Out-Null
                Write-Output ("svg: {0}" -f $svgFile)
            } catch {
                Write-Warning ("svg export failed on slide {0}: {1}" -f $i, $_.Exception.Message)
            }
        }
    }
}
finally {
    if ($pres) { $pres.Close() }
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
