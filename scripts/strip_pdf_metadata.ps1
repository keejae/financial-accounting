# strip_pdf_metadata.ps1 -- remove personal metadata from published PDFs.
#
# strip_metadata.ps1 handles .docx only. Now that lectures publish as PDF, Word /
# PowerPoint exports sometimes carry an /Author stamp ("Hong, Philip Keejae") in the
# Info dictionary and an <dc:creator> entry in the XMP packet. This clears both.
#
# Byte-level, LENGTH-PRESERVING: every removed character is overwritten with a space so
# the file size never changes and the cross-reference table stays valid. Editing a PDF
# any other way (rewriting the string shorter) shifts every later object and corrupts it.
#
# Only two regions are touched -- the /Author(...) string and the XMP packet. Page
# content streams are never modified, so nothing visible on the page can change.
#
# Usage:
#   ./scripts/strip_pdf_metadata.ps1 -Path files -Recurse
#   ./scripts/strip_pdf_metadata.ps1 -Files @('files\week4\korean\Wk4_L13_notes.pdf')

[CmdletBinding()]
param(
    [string]   $Path,
    [switch]   $Recurse,
    [string[]] $Files
)

# ISO-8859-1: byte <-> char is 1:1, so offsets found in the string are byte offsets.
$enc = [System.Text.Encoding]::GetEncoding(28591)

# Names blanked inside the XMP packet. /Author is cleared wholesale regardless.
$names = @(
    'Hong, Philip Keejae', 'Philip Keejae Hong', 'Philip K. Hong', 'Philip Hong',
    'Keejae Hong', 'Hong Philip', 'keejae', 'hong1p', 'cmich'
)

if ($Files) {
    $targets = $Files | ForEach-Object { Get-Item -LiteralPath $_ -ErrorAction SilentlyContinue }
} elseif ($Path) {
    $gci = @{ LiteralPath = $Path; File = $true }
    if ($Recurse) { $gci.Recurse = $true }
    # NOTE: -Include is silently ignored alongside -LiteralPath, so filter afterwards.
    $targets = Get-ChildItem @gci | Where-Object { $_.Extension -eq '.pdf' }
} else {
    Write-Host 'Specify -Path (optionally with -Recurse) or -Files.' -ForegroundColor Yellow
    return
}

$cleaned = 0
foreach ($f in $targets) {
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $text  = $enc.GetString($bytes)
    $hits  = @()

    # ---- 1. Info dictionary: /Author ( ... ) -- blank the whole string ----
    foreach ($m in [regex]::Matches($text, '/Author\s*\((?:\\.|[^)\\])*\)')) {
        $inner = $m.Index + $m.Value.IndexOf('(') + 1
        $len   = $m.Value.LastIndexOf(')') - $m.Value.IndexOf('(') - 1
        if ($len -gt 0) { $hits += ,@($inner, $len) }
    }

    # ---- 2. XMP packet: blank any known name inside it ----
    foreach ($pk in [regex]::Matches($text, '(?s)<\?xpacket begin.*?<\?xpacket end[^>]*>')) {
        foreach ($n in $names) {
            foreach ($m in [regex]::Matches($pk.Value, [regex]::Escape($n), 'IgnoreCase')) {
                $hits += ,@(($pk.Index + $m.Index), $m.Length)
            }
        }
    }

    if ($hits.Count -eq 0) { continue }
    foreach ($h in $hits) {
        for ($i = $h[0]; $i -lt $h[0] + $h[1]; $i++) { $bytes[$i] = 0x20 }
    }
    [System.IO.File]::WriteAllBytes($f.FullName, $bytes)
    $cleaned++
    Write-Host ("  cleaned  {0}" -f $f.Name) -ForegroundColor Green
}

Write-Host ("Scanned {0} PDF(s); cleaned {1}." -f @($targets).Count, $cleaned)
