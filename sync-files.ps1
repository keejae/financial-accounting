# =====================================================================
#  sync-files.ps1
#  Copies lecture files from the Dropbox teaching-notes folder into the
#  website's files/ folder under STABLE, permanent names.
#
#  Why: the website's download links point at the stable names below
#  (e.g. Wk1_L3_slides.pptx). When you revise a lecture, just keep the
#  SAME Dropbox file name, re-run this script, then commit & push.
#  The hyperlinks on the site never have to change.
#
#  HOW TO RUN (from the repo root, in PowerShell):
#     ./sync-files.ps1
#
#  If a source file is renamed in Dropbox, update its name in the
#  $map table below (left-hand side only).
# =====================================================================

$ErrorActionPreference = 'Stop'

# --- Source (Dropbox) and destination (repo) roots ---
$src  = 'C:\Users\hong1p\Dropbox\1_0_teaching\0_2026_Summer\0_Kaist MBA\1_weekly notes'
$repo = Join-Path $PSScriptRoot 'files'

# --- Mapping: source file name  =>  destination path (relative to files/) ---
# Order: Week / Lecture / language.  english = English tab, korean = Korean tab.
$map = [ordered]@{

  # ---------------- WEEK 1 ----------------
  # Lecture 1  (Ch 0 - Color Accounting 1 & 2)  -- two Korean-only slide decks (Introduction + Classic Transactions), reused for English
  '1_Introduction to Accounting - Color Accounting (Korean only) - Students deploy.pptx'    = 'week1\english\Wk1_L1_slides_pt1.pptx'
  '1_Introduction to Accounting - Color Accounting (Korean only) - Students deploy.pptx#ko'  = 'week1\korean\Wk1_L1_slides_pt1.pptx'
  '2_1_Classic Transactions Color Accounting KAIST( Korean only)  Students deploy.pptx'      = 'week1\english\Wk1_L1_slides_pt2.pptx'
  '2_1_Classic Transactions Color Accounting KAIST( Korean only)  Students deploy.pptx#ko'    = 'week1\korean\Wk1_L1_slides_pt2.pptx'
  '1_to_2_Activity Book Color Accounting (Ch 0  English) - Students.docx'  = 'week1\english\Wk1_L1_notes.docx'
  '1_to_2_Activity Book Color Accounting (Ch 0 Korean) - Students.docx'    = 'week1\korean\Wk1_L1_notes.docx'

  # Lecture 2  (Ch 0 - Color Accounting 2 & 3, Journals & T-Accounts)  -- single Korean-only deck, reused for English.
  # (Activity Book Ch 0 lives with Lecture 1 only -- it is identical for both sessions.)
  '2_2_Journals and T-Accounts KAIST (Korean only) Students deploy.pptx'    = 'week1\english\Wk1_L2_slides.pptx'
  '2_2_Journals and T-Accounts KAIST (Korean only) Students deploy.pptx#ko'  = 'week1\korean\Wk1_L2_slides.pptx'

  # Lecture 3  (Ch 1 - Introducing Financial Accounting)
  '3_Introducting Financial Accounting  (Ch 1 English) KAIST - Professor.pptx'          = 'week1\english\Wk1_L3_slides.pptx'
  '3_Introducing_Financial_Accounting (Ch 1 Korean) KAIST - Students deploy.pptx'       = 'week1\korean\Wk1_L3_slides.pptx'
  '3_Introducing Financial Accounting  (Ch 1  English)- Students.docx'                  = 'week1\english\Wk1_L3_notes.docx'
  '3_Introducing Financial Accounting  (Ch 1  Korean)- Students.docx'                   = 'week1\korean\Wk1_L3_notes.docx'

  # Lecture 4  (Ch 2 - Constructing Financial Statements)
  '4_Constructing Financial Statements (Ch 2 English) KAIST- Professor.pptx'            = 'week1\english\Wk1_L4_slides.pptx'
  '4_Constructing_Financial_Statements (Ch 2 Korean) KAIST- Students deploy.pptx'       = 'week1\korean\Wk1_L4_slides.pptx'
  '4_Constructing Financial Statements (Ch 2 English) - Students.docx'                  = 'week1\english\Wk1_L4_notes.docx'
  '4_Constructing Financial Statements (Ch 2 Korean) - Students.docx'                   = 'week1\korean\Wk1_L4_notes.docx'

  # ---------------- WEEK 2 ----------------
  # Lecture 5  (Ch 3 - Adjusting Accounts for Financial Statements)
  '5_Adjusting Accounts for Financial Statements (Ch 3 English) KAIST - Professor.pptx'       = 'week2\english\Wk2_L5_slides.pptx'
  '5_Adjusting_Accounts_for_Financial_Statements (Ch_3_Korean)  KAIST - Students deploy.pptx' = 'week2\korean\Wk2_L5_slides.pptx'
  '5_Adjusting Accounts for Financial Statements (Ch 3 English) - Students.docx'              = 'week2\english\Wk2_L5_notes.docx'
  '5_Adjusting Accounts for Financial Statements (Ch 3 Korean) - Students.docx'               = 'week2\korean\Wk2_L5_notes.docx'

  # Lecture 6  (Ch 4 - Reporting and Analyzing Cash Flows)
  '6_Statement of Cash Flows (Ch 4 English) KAIST - Professor.pptx'                     = 'week2\english\Wk2_L6_slides.pptx'
  '6_Statement of Cash Flows (Ch 4 Korean) KAIST - Students deploy.pptx'                = 'week2\korean\Wk2_L6_slides.pptx'
  '6_Statement of Cash Flows short (Ch 4 English) -  Students.docx'                     = 'week2\english\Wk2_L6_notes.docx'
  '6_Statement of Cash Flows short (Ch 4 Korean) - Students.docx'                       = 'week2\korean\Wk2_L6_notes.docx'

  # Lecture 7  (Ch 6 - Revenue, Receivables & Operating Income)
  '7_Revenue_and_Receivables (Ch 6 English) KAIST Professor.pptx'                       = 'week2\english\Wk2_L7_slides.pptx'
  '7_Revenue_and_Receivables (Ch 6 Korean) KAIST Students deploy.pptx'                  = 'week2\korean\Wk2_L7_slides.pptx'
  '7_Revenue and Account Receivables (Ch 6 English)  shorter version - Students.docx'   = 'week2\english\Wk2_L7_notes.docx'
  '7_Revenue and Account Receivables (Ch 6 Korean) shorter version - Students.docx'     = 'week2\korean\Wk2_L7_notes.docx'

  # Lecture 8  (Ch 7 - Reporting and Analyzing Inventory)
  '8_Inventory (Ch 7 English) KAIS ProfessorT.pptx'                                     = 'week2\english\Wk2_L8_slides.pptx'
  '8_Inventory (Ch 7 Korean) KAIST -Students deploy.pptx'                               = 'week2\korean\Wk2_L8_slides.pptx'
  '8_Reporting and Analyzing Inventory (Ch 7 English) - Students.docx'                  = 'week2\english\Wk2_L8_notes.docx'
  '8_Reporting and Analyzing Inventory (Ch 7 Korean) - Students.docx'                   = 'week2\korean\Wk2_L8_notes.docx'

  # ---------------- WEEK 3 ----------------
  # Lecture 9  (Ch 5 - Analyzing and Interpreting Financial Statements)
  '9_Ratio_Analysis (Ch 5 English) KAIST - Professor.pptx'                              = 'week3\english\Wk3_L9_slides.pptx'
  '9_Ratio_Analysis (Ch_5 Korean) KAIST - Students deploy.pptx'                         = 'week3\korean\Wk3_L9_slides.pptx'
  '9_Analyzing and Interpreting Financial Statements (Ch 5 English) - Students.docx'    = 'week3\english\Wk3_L9_notes.docx'
  '9_Analyzing and Interpreting Financial Statements (Ch 5 Korean) - Students.docx'     = 'week3\korean\Wk3_L9_notes.docx'
}

$copied  = 0
$missing = @()

foreach ($entry in $map.GetEnumerator()) {
    # Strip the disambiguating '#...' suffix used to reuse one source for two destinations.
    $srcName = ($entry.Key -split '#')[0]
    $srcPath = Join-Path $src $srcName
    $dstPath = Join-Path $repo $entry.Value

    if (-not (Test-Path -LiteralPath $srcPath)) {
        $missing += $srcName
        continue
    }

    $dstDir = Split-Path -Parent $dstPath
    if (-not (Test-Path -LiteralPath $dstDir)) {
        New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
    }

    Copy-Item -LiteralPath $srcPath -Destination $dstPath -Force
    $copied++
    Write-Host ("  OK  {0}" -f $entry.Value)
}

Write-Host ""
Write-Host ("Copied $copied file(s).")
if ($missing.Count -gt 0) {
    Write-Host ""
    Write-Host "MISSING SOURCE FILES (not found in Dropbox folder):" -ForegroundColor Yellow
    $missing | Sort-Object -Unique | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Yellow }
}
