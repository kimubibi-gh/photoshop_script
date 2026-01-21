#Requires -Version 5.1
param(
  [Parameter(ValueFromRemainingArguments = $true)]
  [string[]]$Folders = $args
)

# ===== 設定 =====
# 並び方: "Name"（単純な名前昇順） / "Explorer"（数字を自然順に扱うナチュラルソート）
$SortMode = 'Name'

# リネーム後に Photoshop を起動して JSX を走らせるか
$EnablePhotoshopLaunch = $true

# JSX の候補（PS1の隣 or i\ 配下）
$JsxCandidates = @(
  (Join-Path $PSScriptRoot 'i\square_pack.jsx'),
  (Join-Path $PSScriptRoot 'square_pack.jsx')
)

# Photoshop 実行ファイル候補
$PhotoshopCandidates = @(
  'C:\Program Files\Adobe\Adobe Photoshop 2025\Photoshop.exe',
  'C:\Program Files\Adobe\Adobe Photoshop 2024\Photoshop.exe'
)

# JSX に渡す OUT_DIR の方針："Same" or "Subfolder"
$OutDirMode = 'Same'
# ===============

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-BaseFromFolder([string]$FolderPath){
  $leaf = Split-Path -Leaf $FolderPath
  $leafTrim = $leaf.TrimEnd([char]32, [char]0x3000)
  $lastAscii   = $leafTrim.LastIndexOf([char]32)
  $lastZenkaku = $leafTrim.LastIndexOf([char]0x3000)
  $last = [Math]::Max($lastAscii, $lastZenkaku)
  $base = if ($last -ge 0 -and $last -lt ($leafTrim.Length - 1)) { $leafTrim.Substring($last + 1) } else { $leafTrim }
  $invalid = [IO.Path]::GetInvalidFileNameChars()
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $base.ToCharArray()) {
    if ($invalid -notcontains $ch) { [void]$sb.Append($ch) } else { [void]$sb.Append('_') }
  }
  ($sb.ToString().Trim())
}

function ToLowerNormalizedExt([string]$FileName){
  $ext = [IO.Path]::GetExtension($FileName)
  if ([string]::IsNullOrEmpty($ext)) { return "" }
  $ext = $ext.Substring(1).ToLowerInvariant()
  if ($ext -eq 'jpeg') { $ext = 'jpg' }
  return $ext
}

# --- PS5互換：ファイルの並び取得 ---
function Get-FilesSorted([string]$dir, [string]$mode){
  $items = @(Get-ChildItem -LiteralPath $dir -File -Force)
  if ($items.Length -le 1) { return $items }

  if ($mode -eq 'Explorer') {
    # エクスプローラ風ナチュラルソート（StrCmpLogicalW）を自前実装
    Add-Type -ErrorAction SilentlyContinue @"
using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.IO;
public class FileInfoLogicalComparer : IComparer {
  [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
  public static extern int StrCmpLogicalW(string x, string y);
  public int Compare(object a, object b){
    var fa = a as FileSystemInfo;
    var fb = b as FileSystemInfo;
    string sa = (fa != null) ? fa.Name : (a == null ? null : a.ToString());
    string sb = (fb != null) ? fb.Name : (b == null ? null : b.ToString());
    return StrCmpLogicalW(sa, sb);
  }
}
"@
    $list = New-Object System.Collections.ArrayList
    foreach($f in $items){ [void]$list.Add($f) }
    $cmp = New-Object FileInfoLogicalComparer
    $list.Sort($cmp)
    return ,@($list)
  }

  # 単純な名前昇順（PS5互換）
  return @($items | Sort-Object -Property Name)
}

function GetPhotoshopExe(){
  foreach($p in $PhotoshopCandidates){ if (Test-Path -LiteralPath $p) { return $p } }
  return 'Photoshop.exe'
}
function GetMainJsx(){
  foreach($j in $JsxCandidates){ if (Test-Path -LiteralPath $j) { return (Resolve-Path -LiteralPath $j).Path } }
  return $null
}

if (-not $Folders -or $Folders.Count -eq 0) {
  Write-Host "フォルダを引数に渡してください。（D&D可）"
  exit 1
}

foreach ($raw in $Folders) {
  try { $folder = (Resolve-Path -LiteralPath $raw).Path }
  catch { Write-Warning ("Not found: {0}" -f $raw); continue }
  if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
    Write-Warning ("Not a folder: {0}" -f $folder); continue
  }

  Write-Host "[Start] $folder"

  $base = Get-BaseFromFolder $folder
  if ([string]::IsNullOrWhiteSpace($base)) {
    Write-Warning ("Empty base for folder: {0}. Skip." -f (Split-Path -Leaf $folder))
    continue
  }

  # ★ 名前昇順（既定）/Explorerモードはナチュラルソート
  $files = Get-FilesSorted -dir $folder -mode $SortMode
  if ($files.Length -eq 0) { Write-Host ("  No files: {0}" -f $folder); continue }

  # 一時退避（拡張子正規化）
  $map = @()
  $i = 1
  foreach ($f in $files) {
    $extLower = ToLowerNormalizedExt $f.Name
    $extPart  = if ($extLower) { '.' + $extLower } else { '' }
    $tmpLeaf  = ('__tmp_{0}_{1}{2}' -f ([guid]::NewGuid().ToString('N')), $i, $extPart)
    Rename-Item -LiteralPath $f.FullName -NewName $tmpLeaf
    $map += [pscustomobject]@{
      Tmp      = (Join-Path $folder $tmpLeaf)
      Final    = if ($extLower) { ('{0}-{1}.{2}' -f $base, $i, $extLower) } else { ('{0}-{1}' -f $base, $i) }
      ExtLower = $extLower
    }
    $i++
  }

  # 衝突回避しつつ最終名に
  foreach ($m in $map) {
    $finalPath = Join-Path $folder $m.Final
    if (Test-Path -LiteralPath $finalPath) {
      $n = 1
      $nameNoExt = [IO.Path]::GetFileNameWithoutExtension($m.Final)
      $extPart   = [IO.Path]::GetExtension($m.Final)
      do {
        $alt = ('{0}-{1}{2}' -f $nameNoExt, $n, $extPart)
        $finalPath = Join-Path $folder $alt
        $n++
      } while (Test-Path -LiteralPath $finalPath)
    }
    Rename-Item -LiteralPath $m.Tmp -NewName ([IO.Path]::GetFileName($finalPath))
  }

  Write-Host ("  Renamed: {0} file(s). Base='{1}'" -f $map.Count, $base)

  # ===== Photoshop 起動（任意）=====
  if ($EnablePhotoshopLaunch) {
    $mainJsx = GetMainJsx
    if (-not $mainJsx) {
      Write-Warning "square_pack.jsx が見つからないため Photoshop 実行をスキップします。（PS1隣 or i\ 配下）"
    } else {
      $psExe = GetPhotoshopExe
      $inDir = $folder
      switch ($OutDirMode) {
        'Subfolder' {
          $outDir = Join-Path $folder ("加工済み_{0}" -f $base)
          if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
        }
        default { $outDir = $folder }
      }
      $runnerPath = Join-Path $env:TEMP ("ps_runner_{0}.jsx" -f ([guid]::NewGuid().ToString('N')))
      $escIn   = ($inDir   -replace '\\','/')
      $escOut  = ($outDir  -replace '\\','/')
      $escMain = ($mainJsx -replace '\\','/')

      $runner = @"
#target photoshop
app.displayDialogs = DialogModes.NO;
var IN_DIR  = "$escIn";
var OUT_DIR = "$escOut";
var MAIN_JSX = new File("$escMain");
if (!MAIN_JSX.exists) { throw new Error("MAIN JSX not found: " + MAIN_JSX.fsName); }
$.evalFile(MAIN_JSX);
if (typeof run === "function") { run(IN_DIR, OUT_DIR); }
"@
      Set-Content -LiteralPath $runnerPath -Value $runner -Encoding UTF8
      try {
        try { Start-Process -FilePath $psExe -ArgumentList @('/r',"`"$runnerPath`"") -Wait }
        catch { Start-Process -FilePath $psExe -ArgumentList @('-r',"`"$runnerPath`"") -Wait }
        Write-Host ("  Photoshop ran: IN='{0}'  OUT='{1}'" -f $inDir, $outDir)
      } finally {
        Remove-Item -LiteralPath $runnerPath -Force -ErrorAction SilentlyContinue
      }
    }
  }

  Write-Host "[Finish] $folder"
}
