# =====================================================================
#  strip_metadata.ps1
#  Removes personal metadata from Office files before they are published.
#
#  Clears, in each file's zip XML:
#     docProps/core.xml : dc:title, dc:creator, cp:lastModifiedBy
#     docProps/app.xml  : Company
#  and, for .docx only, injects <w:removePersonalInformation/> into
#  word/settings.xml -- that IS Word's "Remove personal information from
#  file properties on save", so Word will not re-stamp the name later.
#  (PowerPoint has no per-file equivalent; .pptx must be re-checked by hand
#  after any save.)
#
#  HOW TO RUN (from the repo root, in PowerShell):
#     ./scripts/strip_metadata.ps1 -Path 'C:\...\HW'            # whole folder
#     ./scripts/strip_metadata.ps1 -Path 'C:\...\HW' -Recurse
#     ./scripts/strip_metadata.ps1 -Files @('a.docx','b.pptx')  # specific files
#
#  Run this on the DROPBOX SOURCE files, then re-run sync-files.ps1 so the
#  site serves the cleaned copies.
# =====================================================================

[CmdletBinding()]
param(
    [string]   $Path,
    [string[]] $Files,
    [switch]   $Recurse,
    # Sub-folder names to skip when scanning a folder (e.g. archived drafts).
    [string[]] $Exclude = @('morgue')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---- Build the work list -------------------------------------------------
$targets = @()
if ($Files) { $targets += $Files | ForEach-Object { (Resolve-Path -LiteralPath $_).Path } }
if ($Path) {
    $gci = @{ LiteralPath = $Path; File = $true; Include = @('*.docx','*.pptx','*.xlsx') }
    if ($Recurse) { $gci.Recurse = $true }
    $targets += Get-ChildItem @gci -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '~$*' } |   # skip Word/PowerPoint lock files
        Where-Object {
            # Skip anything sitting inside an excluded sub-folder (e.g. morgue).
            $parts = $_.FullName -split '\\'
            -not ($Exclude | Where-Object { $parts -contains $_ })
        } | ForEach-Object { $_.FullName }
}
$targets = $targets | Sort-Object -Unique
if (-not $targets) { Write-Host 'No files matched.' -ForegroundColor Yellow; return }

# ---- Helpers -------------------------------------------------------------
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Get-EntryText($zip, $name) {
    $e = $zip.Entries | Where-Object { $_.FullName -eq $name }
    if (-not $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open())
    $t = $sr.ReadToEnd(); $sr.Close()
    return $t
}

function Set-EntryText($zip, $name, $text) {
    # Delete + recreate: ZipArchiveEntry has no in-place resize.
    $e = $zip.Entries | Where-Object { $_.FullName -eq $name }
    if ($e) { $e.Delete() }
    $new = $zip.CreateEntry($name)
    $st = $new.Open()
    $bytes = $utf8NoBom.GetBytes($text)
    $st.Write($bytes, 0, $bytes.Length)
    $st.Close()
}

# ---- Clean ---------------------------------------------------------------
$cleaned = 0
foreach ($file in $targets) {
    $changed = $false
    $zip = [System.IO.Compression.ZipFile]::Open($file, 'Update')
    try {
        # core.xml -- title, author, last-modified-by
        $core = Get-EntryText $zip 'docProps/core.xml'
        if ($core) {
            $new = $core
            foreach ($tag in @('dc:title','dc:creator','cp:lastModifiedBy')) {
                $new = [regex]::Replace($new, "<$tag>.*?</$tag>", "<$tag></$tag>")
            }
            if ($new -ne $core) { Set-EntryText $zip 'docProps/core.xml' $new; $changed = $true }
        }

        # app.xml -- company
        $app = Get-EntryText $zip 'docProps/app.xml'
        if ($app) {
            $new = [regex]::Replace($app, '<Company>.*?</Company>', '<Company></Company>')
            if ($new -ne $app) { Set-EntryText $zip 'docProps/app.xml' $new; $changed = $true }
        }

        # settings.xml (.docx only) -- keep it removed on future saves
        if ([System.IO.Path]::GetExtension($file) -eq '.docx') {
            $settings = Get-EntryText $zip 'word/settings.xml'
            if ($settings -and $settings -notmatch 'removePersonalInformation') {
                # Must sit at the start of the settings element to satisfy the schema.
                $new = [regex]::Replace($settings, '(<w:settings\b[^>]*>)',
                                        '$1<w:removePersonalInformation/>', 1)
                if ($new -ne $settings) { Set-EntryText $zip 'word/settings.xml' $new; $changed = $true }
            }
        }
    }
    finally { $zip.Dispose() }

    if ($changed) { $cleaned++; Write-Host ("  CLEANED  {0}" -f (Split-Path -Leaf $file)) }
    else          {           Write-Host ("  ok       {0}" -f (Split-Path -Leaf $file)) -ForegroundColor DarkGray }
}

Write-Host ""
Write-Host ("Cleaned $cleaned of $($targets.Count) file(s).")
