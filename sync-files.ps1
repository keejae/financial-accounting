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
  '2_1_Classic Transactions Color Accounting (Korean_only) - Students deploy.pptx'      = 'week1\english\Wk1_L1_slides_pt2.pptx'
  '2_1_Classic Transactions Color Accounting (Korean_only) - Students deploy.pptx#ko'    = 'week1\korean\Wk1_L1_slides_pt2.pptx'
  '1_to_2_Activity Book Color Accounting (Ch 0  English) - Students.docx'  = 'week1\english\Wk1_L1_notes.docx'
  '1_to_2_Activity Book Color Accounting (Ch 0 Korean) - Students.docx'    = 'week1\korean\Wk1_L1_notes.docx'
  # Professor (worked-answer) copies -- posted only AFTER the week has been taught.
  '1_to_2_Activity Book Color Accounting  (Ch 0 Korean) - Professor.docx'  = 'week1\korean\Wk1_L1_notes_prof.docx'

  # Lecture 2  (Ch 0 - Color Accounting 2 & 3, Journals & T-Accounts)  -- single Korean-only deck, reused for English.
  # (Activity Book Ch 0 lives with Lecture 1 only -- it is identical for both sessions.)
  '2_2_Journals and T-Accounts KAIST (Korean only) Students deploy.pptx'    = 'week1\english\Wk1_L2_slides.pptx'
  '2_2_Journals and T-Accounts KAIST (Korean only) Students deploy.pptx#ko'  = 'week1\korean\Wk1_L2_slides.pptx'

  # Lecture 3  (Ch 1 - Introducing Financial Accounting)
  '3_Introducting Financial Accounting  (Ch 1 English) KAIST - Professor.pptx'          = 'week1\english\Wk1_L3_slides.pptx'
  '3_Introducing_Financial_Accounting (Ch 1 Korean_f2) KAIST - Students deploy.pptx'    = 'week1\korean\Wk1_L3_slides.pptx'
  '3_Introducing Financial Accounting  (Ch 1  English)- Students.docx'                  = 'week1\english\Wk1_L3_notes.docx'
  '3_Introducing Financial Accounting  (Ch 1  Korean)- Students.docx'                   = 'week1\korean\Wk1_L3_notes.docx'
  '3_Introducing Financial Accounting  (Ch 1  Korean)-  Professor.docx'                 = 'week1\korean\Wk1_L3_notes_prof.docx'

  # Lecture 4  (Ch 2 - Constructing Financial Statements)
  '4_Constructing Financial Statements (Ch 2 English) KAIST- Professor.pptx'            = 'week1\english\Wk1_L4_slides.pptx'
  '4_Constructing_Financial_Statements (Ch 2 Korean_f2) KAIST- Students deploy.pptx'    = 'week1\korean\Wk1_L4_slides.pptx'
  '4_Constructing Financial Statements (Ch 2 English) - Students.docx'                  = 'week1\english\Wk1_L4_notes.docx'
  '4_Constructing Financial Statements (Ch 2 Korean) - Students.docx'                   = 'week1\korean\Wk1_L4_notes.docx'
  '4_Constructing Financial Statements (Ch 2 Korean) - Professor.docx'                  = 'week1\korean\Wk1_L4_notes_prof.docx'

  # ---------------- WEEK 2 ----------------
  # Lecture 5  (Ch 3 - Adjusting Accounts for Financial Statements)
  '5_Adjusting Accounts for Financial Statements (Ch 3 English) KAIST - Professor.pptx'       = 'week2\english\Wk2_L5_slides.pptx'
  # 2026-08-03: repointed off the _f2 file to the deck Philip revised directly.
  '5_Adjusting_Accounts_for_Financial_Statements (Ch_3_Korean)  KAIST - Students deploy.pptx'    = 'week2\korean\Wk2_L5_slides.pptx'
  '5_Adjusting Accounts for Financial Statements (Ch 3 English) - Students.docx'              = 'week2\english\Wk2_L5_notes.docx'
  '5_Adjusting Accounts for Financial Statements (Ch 3 Korean) - Students.docx'               = 'week2\korean\Wk2_L5_notes.docx'

  # Lecture 6  (Ch 6 - Revenue, Receivables & Operating Income)
  # 2026-08-05 schedule revision: Ch 6 moved from Lecture 7 to Lecture 6 (swapped with Ch 4).
  '7_Revenue_and_Receivables (Ch 6 English) KAIST Professor.pptx'                       = 'week2\english\Wk2_L6_slides.pptx'
  '7_Revenue_and_Receivables (Ch 6 Korean_f2) KAIST Students deploy.pptx'               = 'week2\korean\Wk2_L6_slides.pptx'
  '7_Revenue and Account Receivables (Ch 6 English)  shorter version - Students.docx'   = 'week2\english\Wk2_L6_notes.docx'
  '7_Revenue and Account Receivables (Ch 6 Korean) shorter version - Students.docx'     = 'week2\korean\Wk2_L6_notes.docx'

  # Lecture 7  (Ch 4 - Reporting and Analyzing Cash Flows)
  '6_Statement of Cash Flows (Ch 4 English) KAIST - Professor.pptx'                     = 'week2\english\Wk2_L7_slides.pptx'
  '6_Statement of Cash Flows (Ch 4 Korean_f2) KAIST - Students deploy.pptx'             = 'week2\korean\Wk2_L7_slides.pptx'
  '6_Statement of Cash Flows short (Ch 4 English) -  Students.docx'                     = 'week2\english\Wk2_L7_notes.docx'
  '6_Statement of Cash Flows short (Ch 4 Korean) - Students.docx'                       = 'week2\korean\Wk2_L7_notes.docx'

  # Lecture 8  (Special Topic - Valuation Basics and P/E Ratio)
  # Not yet posted -- Philip is still preparing this session. Add the mapping when ready.

  # ---------------- WEEK 3 ----------------
  # Lecture 9  (Ch 7 - Reporting and Analyzing Inventory)   -- was Week 2 Lecture 8 before the 2026-08-05 revision
  '8_Inventory (Ch 7 English) KAIS ProfessorT.pptx'                                     = 'week3\english\Wk3_L9_slides.pptx'
  '8_Inventory (Ch 7 Korean_f2) KAIST -Students deploy.pptx'                            = 'week3\korean\Wk3_L9_slides.pptx'
  '8_Reporting and Analyzing Inventory (Ch 7 English) - Students.docx'                  = 'week3\english\Wk3_L9_notes.docx'
  '8_Reporting and Analyzing Inventory (Ch 7 Korean) - Students.docx'                   = 'week3\korean\Wk3_L9_notes.docx'

  # Lecture 10 (Ch 5 - Analyzing and Interpreting Financial Statements)  -- was Lecture 9
  '9_Ratio_Analysis (Ch 5 English) KAIST - Professor.pptx'                              = 'week3\english\Wk3_L10_slides.pptx'
  '9_Ratio_Analysis (Ch_5 Korean_f2) KAIST - Students deploy.pptx'                      = 'week3\korean\Wk3_L10_slides.pptx'
  '9_Analyzing and Interpreting Financial Statements (Ch 5 English) - Students.docx'    = 'week3\english\Wk3_L10_notes.docx'
  '9_Analyzing and Interpreting Financial Statements (Ch 5 Korean) - Students.docx'     = 'week3\korean\Wk3_L10_notes.docx'

  # ---------------- IN-CLASS ACTIVITIES (weeks/inclass.html) ----------------
  # Korean only for now; add an 'inclass\english\...' line when an English version exists.
  # Destination lecture numbers follow the CURRENT schedule, source prefixes the original
  # one -- match on the CHAPTER (the '7_' file is Ch 6, which is now Lecture 6).
  # STUDENT copies only -- never map the '... InClass Professor Copy.docx' files.
  '5_InClass Questions Adjusting Accounts for Financial Statements (Ch 3 Korean).docx'  = 'inclass\korean\IC_L5_activities.docx'
  '7_Revenue and Account Receivables - InClass Problems.pdf'                            = 'inclass\korean\IC_L6_activities.pdf'

  # ---------------- HOMEWORK (weeks/hw.html, weeks/hw-solutions.html) ----------------
  # Sources live in the 'HW' subfolder. Its numeric file prefixes follow the ORIGINAL
  # lecture numbering; the stable destination names follow the CURRENT (2026-08-05)
  # schedule, so the two no longer line up -- match on the CHAPTER, not the prefix:
  #     5_  = Ch 3 -> Lecture 5      7_  = Ch 6 -> Lecture 6
  #     8_  = Ch 7 -> Lecture 9      9_  = Ch 5 -> Lecture 10     10_ = Ch 8 -> Lecture 13
  # (Lecture 7 / Ch 4 has no homework.)
  # Solutions go up in the same week the lecture is taught, so add a '_solutions' line
  # per assignment as that week arrives -- not before.
  'HW\5_HW Adjusting Accounts for Financial Statements (English Ch 3) - Problems.docx' = 'hw\english\HW_L5_problems.docx'
  'HW\5_HW Adjusting Accounts for Financial Statements (Korean Ch 3) - Problems.docx'  = 'hw\korean\HW_L5_problems.docx'
  'HW\5_HW Adjusting Accounts for Financial Statements (English Ch 3) - Solutions.docx' = 'hw\english\HW_L5_solutions.docx'
  # NOTE: the Korean Ch 3 solutions file is spelled "Soutions" in Dropbox -- keep as is.
  'HW\5_HW Adjusting Accounts for Financial Statements (Korean Ch 3) - Soutions.docx'  = 'hw\korean\HW_L5_solutions.docx'
  'HW\7_HW Revenue Recognition and Acct Rec (Engslish Ch 6) - Problems.docx'           = 'hw\english\HW_L6_problems.docx'
  'HW\7_HW Revenue Recognition and Acct Rec (Korean Ch 6) - Problems.docx'             = 'hw\korean\HW_L6_problems.docx'
  'HW\7_HW Revenue Recognition and Acct Rec (Engslish Ch 6) - Solutions.docx'          = 'hw\english\HW_L6_solutions.docx'
  'HW\7_HW Revenue Recognition and Acct Rec (Korean Ch 6) - Solutions.docx'            = 'hw\korean\HW_L6_solutions.docx'
  'HW\8_HW Inventory (Engslish Ch 7) - Problems.docx'                                  = 'hw\english\HW_L9_problems.docx'
  'HW\8_HW Inventory (Korean Ch 7) - Problems.docx'                                    = 'hw\korean\HW_L9_problems.docx'
  'HW\9_HW Ratio Analysis  (English Ch 5) - Problems.docx'                             = 'hw\english\HW_L10_problems.docx'
  'HW\9_HW Ratio Analysis  (Korean Ch 5) - Problems.docx'                              = 'hw\korean\HW_L10_problems.docx'
  'HW\10_HW Long-Lived  Assets (English Ch 8) - Problems.docx'                         = 'hw\english\HW_L13_problems.docx'
  'HW\10_HW Long-Lived  Assets (Korean Ch 8) - Problems.docx'                          = 'hw\korean\HW_L13_problems.docx'

  # (Week 2 Lecture 8 - Valuation special topic - and its in-class files are handled
  #  by the runtime block further down: their names contain Hangul.)

  # ---------------- SPECIAL TOPIC 1 (Introduction to Valuation) ----------------
  # REMOVED 2026-08-05 at Philip's request -- the page, its files and the Korean
  # preview video were all taken down. The Dropbox sources
  # ('99_Valuation_Primer_Fin (English|Korean) - Students.docx' and the 99 preview
  # mp4) are untouched, so restoring it is just a matter of re-adding these lines.

  # ---------------- SPECIAL TOPIC 2 (Blockchain) ----------------
  # Bilingual: student slide deck + lesson-plan notes (English & Korean).
  '99_blockchain - PPT Slides (English).pptx'                              = 'special2\english\ST2_Blockchain_slides.pptx'
  '99_blockchain - PPT Slides (Korean).pptx'                               = 'special2\korean\ST2_Blockchain_slides.pptx'
  '99_Exploring Blockchain and Crytpoassets - Lesson Plans (English).docx' = 'special2\english\ST2_Blockchain_notes.docx'
  '99_Exploring Blockchain and Crytpoassets - Lesson Plans (Korean).docx'  = 'special2\korean\ST2_Blockchain_notes.docx'
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

# ---------------- COURSE PREVIEW VIDEOS (Korean only) ----------------
# NotebookLM / Gemini preview videos (recorded in Korean). Matched by numeric
# prefix so we never embed Korean file names in this ANSI-read script.
$videoSrc = Join-Path $src 'video'
$videoMap = [ordered]@{
  '1_*.mp4'  = 'week1\korean\Wk1_L1_preview_pt1.mp4'
  '2_1*.mp4' = 'week1\korean\Wk1_L1_preview_pt2.mp4'
  '2_2*.mp4' = 'week1\korean\Wk1_L2_preview.mp4'
  '3_*.mp4'  = 'week1\korean\Wk1_L3_preview.mp4'
  '4_*.mp4'  = 'week1\korean\Wk1_L4_preview.mp4'
  # 2026-08-05: prefixes still follow the ORIGINAL lecture numbering, destinations follow
  # the CURRENT schedule -- 7_ (Ch 6) is now Lecture 6, 6_ (Ch 4) is Lecture 7, 8_ (Ch 7)
  # moved to Week 3 Lecture 9.
  '5_*.mp4'  = 'week2\korean\Wk2_L5_preview.mp4'
  '7_*.mp4'  = 'week2\korean\Wk2_L6_preview.mp4'
  '6_*.mp4'  = 'week2\korean\Wk2_L7_preview.mp4'
  '8_*.mp4'  = 'week3\korean\Wk3_L9_preview.mp4'
  # '99*.mp4' (Valuation preview) removed 2026-08-05 with the Special Topic 1 page.
}
if (Test-Path -LiteralPath $videoSrc) {
    foreach ($entry in $videoMap.GetEnumerator()) {
        $match = @(Get-ChildItem -LiteralPath $videoSrc -Filter $entry.Key -File -ErrorAction SilentlyContinue)
        if ($match.Count -eq 0) { $missing += ('video\' + $entry.Key); continue }
        if ($match.Count -gt 1) { Write-Host ("  WARN {0} matched {1} files; using first" -f $entry.Key, $match.Count) -ForegroundColor Yellow }
        $dstPath = Join-Path $repo $entry.Value
        $dstDir  = Split-Path -Parent $dstPath
        if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }
        Copy-Item -LiteralPath $match[0].FullName -Destination $dstPath -Force
        $copied++
        Write-Host ("  OK  {0}" -f $entry.Value)
    }
}

# ---------------- WEEK 2 LECTURE 8 + ITS IN-CLASS FILES (Valuation special topic) ----------------
# Sources live in the '0_Valuation' subfolder. Several file names contain Hangul, so
# they are matched at runtime by ASCII wildcard rather than written as literals -- this
# .ps1 is read as ANSI by Windows PowerShell 5.1 and Korean literals would be mangled
# (same reasoning as the video and Week 0 blocks).
# Lecture notes are posted as PDF (not .docx) at Philip's request; Student versions only.
$valSrc = Join-Path $src '0_Valuation'
if (Test-Path -LiteralPath $valSrc) {
    # Some files differ ONLY in Hangul (학생용 = student vs 교수용 = professor), so an ASCII
    # wildcard cannot tell them apart. Build the marker from code points instead -- the
    # script source stays pure ASCII, and the string is correct at runtime.
    $KO_STUDENT = [string]::Concat([char]0xD559, [char]0xC0DD, [char]0xC6A9)   # 학생용

    $valJobs = @(
        # Lecture 8 slides + notes
        @{ Filter = '99_Valuation Primer (English)*Professor*.pptx'; Dest = 'week2\english\Wk2_L8_slides.pptx' }
        @{ Filter = '99_Valuation_Primer_*KAIST-Professor.pptx';     Dest = 'week2\korean\Wk2_L8_slides.pptx' }
        @{ Filter = '99_Valuation_Primer (English) Students.pdf';    Dest = 'week2\english\Wk2_L8_notes.pdf' }
        # Korean notes: same pattern as the English one, so exclude the English match.
        @{ Filter = '99_Valuation_Primer (*) Students.pdf'; Exclude = 'English'
           Dest = 'week2\korean\Wk2_L8_notes.pdf' }
        # In-class activity files for Lecture 8
        @{ Filter = '99_*Memory_Big3_10Y_Key_Metrics*.xlsx'; Dest = 'inclass\korean\IC_L8_memory3_metrics.xlsx' }
        @{ Filter = '99_SKhynix_FCFF_Valuation_Korean.xlsx'; Dest = 'inclass\korean\IC_L8_skhynix_fcff.xlsx' }
        @{ Filter = '99_SKhynix_*.pdf';                      Dest = 'inclass\korean\IC_L8_skhynix_financials.pdf' }
        # Practice problems -- STUDENT copy only. The professor copy sits beside it with an
        # identical ASCII name shape, so it is selected by the Hangul marker above.
        @{ Filter = '99_Valuation_Primer_*.docx'; Include = $KO_STUDENT
           Dest = 'inclass\korean\IC_L8_practice_problems.docx' }
    )
    foreach ($job in $valJobs) {
        $m = @(Get-ChildItem -LiteralPath $valSrc -Filter $job.Filter -File -ErrorAction SilentlyContinue)
        if ($job.Exclude) { $m = @($m | Where-Object { $_.Name -notlike ('*' + $job.Exclude + '*') }) }
        if ($job.Include) { $m = @($m | Where-Object { $_.Name -like     ('*' + $job.Include + '*') }) }
        if ($m.Count -eq 0) { $missing += ('0_Valuation: ' + $job.Filter); continue }
        if ($m.Count -gt 1) { Write-Host ("  WARN {0} matched {1} files; using first" -f $job.Filter, $m.Count) -ForegroundColor Yellow }
        $dstPath = Join-Path $repo $job.Dest
        $dstDir  = Split-Path -Parent $dstPath
        if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }
        Copy-Item -LiteralPath $m[0].FullName -Destination $dstPath -Force
        $copied++
        Write-Host ("  OK  {0}" -f $job.Dest)
    }
}

# ---------------- WEEK 0 (Course Resources: syllabus + accounting terminology) ----------------
# These live in the "Other Resources web" subfolder, split into English / 한국어 language
# subfolders. The Korean folder name and the Korean syllabus file name contain Hangul,
# so we locate them at runtime / by ASCII wildcard to keep this ANSI-read script free of
# non-ASCII literals (same reasoning as the video block above).
$w0src = Join-Path $src 'Other Resources web'
if (Test-Path -LiteralPath $w0src) {
    $w0en    = Join-Path $w0src 'English'
    $w0koDir = Get-ChildItem -LiteralPath $w0src -Directory | Where-Object { $_.Name -ne 'English' } | Select-Object -First 1
    $w0ko    = if ($w0koDir) { $w0koDir.FullName } else { $null }

    $w0jobs = @(
        @{ Dir = $w0en; Filter = '*Syllabus*English*.docx';    Dest = 'week0\english\Wk0_Syllabus.docx' }
        @{ Dir = $w0en; Filter = '*Terminology*English*.xlsx'; Dest = 'week0\english\Wk0_Terminology.xlsx' }
        @{ Dir = $w0ko; Filter = '*Korean*.docx';              Dest = 'week0\korean\Wk0_Syllabus.docx' }
        @{ Dir = $w0ko; Filter = '*Terminology*Korean*.xlsx';  Dest = 'week0\korean\Wk0_Terminology.xlsx' }
    )
    foreach ($job in $w0jobs) {
        if (-not $job.Dir) { $missing += ('week0: ' + $job.Filter); continue }
        $m = @(Get-ChildItem -LiteralPath $job.Dir -Filter $job.Filter -File -ErrorAction SilentlyContinue)
        if ($m.Count -eq 0) { $missing += ('week0: ' + $job.Filter); continue }
        if ($m.Count -gt 1) { Write-Host ("  WARN {0} matched {1} files; using first" -f $job.Filter, $m.Count) -ForegroundColor Yellow }
        $dstPath = Join-Path $repo $job.Dest
        $dstDir  = Split-Path -Parent $dstPath
        if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }
        Copy-Item -LiteralPath $m[0].FullName -Destination $dstPath -Force
        $copied++
        Write-Host ("  OK  {0}" -f $job.Dest)
    }
}

Write-Host ""
Write-Host ("Copied $copied file(s).")
if ($missing.Count -gt 0) {
    Write-Host ""
    Write-Host "MISSING SOURCE FILES (not found in Dropbox folder):" -ForegroundColor Yellow
    $missing | Sort-Object -Unique | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Yellow }
}
