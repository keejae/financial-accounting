param(
  [Parameter(Mandatory=$true)][string[]]$Files,
  [string]$ExportDir = ""   # if set, export slides 1,5,10 of the FIRST file after numbering (for QA)
)

Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'

# --- enums / constants ---
$msoTextHoriz = 1
$ppAlignCenter = 2
$WHITE = 16777215   # 0xFFFFFF
$BLACK = 0
$FONT  = "Arial"
$FONTSIZE = 14
$LUMA_THRESHOLD = 140   # avg luminance below this => dark background => white font

$tmpPng = Join-Path $env:TEMP ("pgnum_sample_{0}.png" -f $PID)

function Get-RegionLuma([string]$png, [int]$x0,[int]$y0,[int]$x1,[int]$y1){
  $bmp = New-Object System.Drawing.Bitmap($png)
  try {
    $x0 = [Math]::Max(0,$x0); $y0 = [Math]::Max(0,$y0)
    $x1 = [Math]::Min($bmp.Width-1,$x1); $y1 = [Math]::Min($bmp.Height-1,$y1)
    $sum = 0.0; $n = 0
    for($x=$x0; $x -le $x1; $x+=1){
      for($y=$y0; $y -le $y1; $y+=1){
        $p = $bmp.GetPixel($x,$y)
        $sum += (0.299*$p.R + 0.587*$p.G + 0.114*$p.B)
        $n++
      }
    }
    if($n -eq 0){ return 255 }
    return $sum / $n
  } finally { $bmp.Dispose() }
}

$pp = New-Object -ComObject PowerPoint.Application
$first = $true
$failed = @()

foreach($file in $Files){
 try {
  Write-Host "Processing: $(Split-Path $file -Leaf)"
  $pres = $pp.Presentations.Open($file, $false, $false, $false)  # editable, no window
  $SW = $pres.PageSetup.SlideWidth
  $SH = $pres.PageSetup.SlideHeight

  # number box geometry (points) -- centered inside the right-edge blue bar, near the bottom
  $boxW = 40; $boxH = 24
  $boxLeft = $SW - 44                        # right edge ~4pt from slide edge
  $boxTop  = $SH - 32
  $cxc = $boxLeft + $boxW/2                   # horizontal centre (where digits render)
  # sample region in SLIDE points, centred where the digits will sit
  $sxL = $cxc - 11; $sxR = $cxc + 11
  $syT = $boxTop + 2; $syB = $boxTop + $boxH - 2

  $expW = 480; $expH = [int][Math]::Round($expW * $SH / $SW)
  $scaleX = $expW / $SW; $scaleY = $expH / $SH

  $count = $pres.Slides.Count
  for($i=1; $i -le $count; $i++){
    $sl = $pres.Slides.Item($i)

    # (a) remove any prior auto page number (idempotent)
    for($k=$sl.Shapes.Count; $k -ge 1; $k--){
      if($sl.Shapes.Item($k).Name -eq "AutoPageNum"){ $sl.Shapes.Item($k).Delete() }
    }

    # (b) export to sample the background behind the number
    $sl.Export($tmpPng, "PNG", $expW, $expH)
    $luma = Get-RegionLuma $tmpPng ([int]($sxL*$scaleX)) ([int]($syT*$scaleY)) ([int]($sxR*$scaleX)) ([int]($syB*$scaleY))
    $color = if($luma -lt $LUMA_THRESHOLD){ $WHITE } else { $BLACK }

    # (c) add the number box
    $box = $sl.Shapes.AddTextbox($msoTextHoriz, $boxLeft, $boxTop, $boxW, $boxH)
    $box.Name = "AutoPageNum"
    $box.Fill.Visible = 0
    $box.Line.Visible = 0
    $tf = $box.TextFrame
    $tf.AutoSize = 0     # ppAutoSizeNone -> keep the fixed box width so centering works
    $tf.WordWrap = -1    # true -> center alignment centres the digit within the box
    $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
    $tr = $tf.TextRange
    $tr.InsertSlideNumber() | Out-Null
    $tr.ParagraphFormat.Alignment = $ppAlignCenter
    $tr.Font.Size = $FONTSIZE
    $tr.Font.Name = $FONT
    $tr.Font.Bold = 0
    $tr.Font.Color.RGB = $color
  }

  # QA export of first file
  if($first -and $ExportDir -ne ""){
    New-Item -ItemType Directory -Force -Path $ExportDir | Out-Null
    foreach($qi in @(1,5,10)){
      if($qi -le $count){ $pres.Slides.Item($qi).Export((Join-Path $ExportDir ("qa_{0}.png" -f $qi)), "PNG", 960, 540) }
    }
  }
  $first = $false

  $pres.Save()
  $pres.Close()
  Write-Host "  done ($count slides)"
 } catch {
  Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
  $failed += (Split-Path $file -Leaf)
  try { if($pres){ $pres.Close() } } catch {}
 }
}

if($failed.Count -gt 0){ Write-Host "`nFAILED FILES:"; $failed | ForEach-Object { Write-Host "  - $_" } }
$pp.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
if(Test-Path $tmpPng){ Remove-Item $tmpPng -Force }
Write-Host "ALL DONE"
