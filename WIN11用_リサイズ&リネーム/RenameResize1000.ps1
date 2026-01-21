param(
  [Parameter(ValueFromRemainingArguments=$true)]
  [string[]]$Folders
)

if (-not $Folders -or $Folders.Count -eq 0) {
  Write-Host "Drag & drop folders onto this script."
  exit 1
}

# settings
$MaxEdge     = 1000   # longest edge (shrink only)
$JpegQuality = 100
$Desktop     = [Environment]::GetFolderPath('Desktop')
$OutRoot     = Join-Path $Desktop 'resized'
$Exts        = @('.jpg','.jpeg')
$ZeroPad     = $false

# load GDI+
try { Add-Type -AssemblyName System.Drawing -ErrorAction Stop }
catch { Write-Error "System.Drawing load failed. Use Windows PowerShell (powershell.exe)."; exit 1 }

function Ensure-Dir([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) { [void](New-Item -ItemType Directory -Path $Path) }
}

# base name = after the last half-width or full-width space
function Get-BaseFromFolder([string]$FolderPath) {
  $leaf = Split-Path $FolderPath -Leaf
  $parts = $leaf -split '[ \u3000]'
  if ($parts.Count -gt 0) { return $parts[$parts.Count-1] } else { return $leaf }
}

# EXIF orientation
function Fix-Orientation([System.Drawing.Image]$img) {
  try {
    $prop = $img.PropertyItems | Where-Object { $_.Id -eq 0x0112 } | Select-Object -First 1
    if ($null -ne $prop) {
      $o = [BitConverter]::ToUInt16($prop.Value,0)
      switch ($o) {
        2 { $img.RotateFlip('RotateNoneFlipX') }
        3 { $img.RotateFlip('Rotate180FlipNone') }
        4 { $img.RotateFlip('RotateNoneFlipY') }
        5 { $img.RotateFlip('Rotate90FlipX') }
        6 { $img.RotateFlip('Rotate90FlipNone') }
        7 { $img.RotateFlip('Rotate270FlipX') }
        8 { $img.RotateFlip('Rotate270FlipNone') }
      }
    }
  } catch {}
}

# jpeg encoder
function Get-JpegEncoder {
  [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
}
$JpegEnc = Get-JpegEncoder

foreach ($rootIn in $Folders) {

  if (-not (Test-Path -LiteralPath $rootIn)) { Write-Warning "Not found: $rootIn"; continue }
  $root = (Resolve-Path -LiteralPath $rootIn).Path
  $leaf = Split-Path $root -Leaf
  $base = Get-BaseFromFolder $root

  Write-Host "[Start] $root"
  Write-Host "  Base = '$base'"

  # pick jpg/jpeg only (folder root only)
  $files = Get-ChildItem -LiteralPath $root -File |
           Where-Object {
             ($Exts -contains ([IO.Path]::GetExtension($_.Name).ToLowerInvariant())) -and
             ($_.Name -notlike '__tmp_*')
           } |
           Sort-Object Name

  if (-not $files) { Write-Host "  [Skip] no jpg/jpeg."; continue }

  # stage1: temp rename (lowercase ext preserved)
  $temps = @()
  foreach ($f in $files) {
    $extLower = [IO.Path]::GetExtension($f.Name).ToLowerInvariant()
    $tmpLeaf  = "__tmp_" + [Guid]::NewGuid().ToString("N") + $extLower
    Rename-Item -LiteralPath $f.FullName -NewName $tmpLeaf
    $temps += (Join-Path $root $tmpLeaf)
  }

  # stage2: final rename starting from 1, ext = .jpg
  $i = 1
  $renamed = @()
  foreach ($tmpPath in $temps) {
    if ($ZeroPad) {
      $num = "{0:D3}" -f $i
    } else {
      $num = "$i"
    }

    $finalLeaf = "$base-$num.jpg"
    $finalPath = Join-Path $root $finalLeaf
    Rename-Item -LiteralPath $tmpPath -NewName $finalLeaf
    $renamed += $finalPath
    $i++
  }
  Write-Host "  Renamed: $($renamed.Count) file(s)."

  # $outBase = Join-Path $OutRoot $base
  $OutRoot = Join-Path ([Environment]::GetFolderPath('Desktop')) 'ƒŠƒTƒCƒYÏ‚İ'
  $outBase = Join-Path $OutRoot ($base)
  Ensure-Dir $outBase

  foreach ($srcPath in ($renamed | Sort-Object)) {
    $dstPath = Join-Path $outBase (Split-Path $srcPath -Leaf)

    if (Test-Path -LiteralPath $dstPath) {
      $bn = [IO.Path]::GetFileNameWithoutExtension($dstPath)
      $n=1
      do {
        $dstPath2 = Join-Path $outBase ("{0}-{1}.jpg" -f $bn,$n)
        $n++
      } while (Test-Path -LiteralPath $dstPath2)
      $dstPath = $dstPath2
    }

    $img=$null; $bmp=$null; $g=$null
    try {
      $img = [System.Drawing.Image]::FromFile($srcPath)
      Fix-Orientation $img

      $w = [int]$img.Width
      $h = [int]$img.Height
      $long = [Math]::Max($w,$h)

      if ($long -le $MaxEdge) {
        Copy-Item -LiteralPath $srcPath -Destination $dstPath
      } else {
        $scale = $MaxEdge / [double]$long
        $nw = [int][Math]::Round($w * $scale)
        $nh = [int][Math]::Round($h * $scale)

        $bmp = New-Object System.Drawing.Bitmap ($nw, $nh)  #     =32bpp
        $g   = [System.Drawing.Graphics]::FromImage($bmp)

        $g.CompositingMode    = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
        $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

        $white = [System.Drawing.Brushes]::White
        $g.FillRectangle($white,0,0,$nw,$nh) | Out-Null
        $g.DrawImage($img,0,0,$nw,$nh) | Out-Null

        $encParams = New-Object System.Drawing.Imaging.EncoderParameters 1
        $encParam  = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality), $JpegQuality
        $encParams.Param[0] = $encParam
        $bmp.Save($dstPath, $JpegEnc, $encParams)
      }
    } catch {
      Write-Host "  ResizeFail: $srcPath  ->  $($_.Exception.Message)"
    } finally {
      if ($g)   { $g.Dispose() }
      if ($bmp) { $bmp.Dispose() }
      if ($img) { $img.Dispose() }
    }
  }

  Write-Host "[Done] $root -> $outBase"
}
