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
  # 2026-08-10: Philip renamed these sources from the '8_' prefix to '9_' to match the current
  # schedule, and now publishes BOTH the slides and the notes as PDF (not .pptx / .docx).
  # The English deck is the Students version as of 2026-08-10 (it used to be the Professor one).
  # NOTE: 'KAIST  Students.pdf' really does carry two spaces -- keep it exact.
  '9_Inventory (Ch 7 English) KAIST  Students.pdf'                                      = 'week3\english\Wk3_L9_slides.pdf'
  '9_Inventory (Ch 7 Korean_f2) KAIST -Students deploy pptx.pdf'                        = 'week3\korean\Wk3_L9_slides.pdf'
  '9_Reporting and Analyzing Inventory (Ch 7 English) - Students.pdf'                   = 'week3\english\Wk3_L9_notes.pdf'
  '9_Reporting and Analyzing Inventory (Ch 7 Korean) - Students.pdf'                    = 'week3\korean\Wk3_L9_notes.pdf'

  # Lecture 10 (Ch 5 - Analyzing and Interpreting Financial Statements)  -- was Lecture 9
  # 2026-08-10: sources renamed from the '9_' prefix to '10_', and this lecture now
  # publishes as PDF too -- students said .pptx / .docx are hard to read on a tablet.
  # STUDENT versions only, both languages. The earlier ambiguity over which Korean deck
  # was current is settled: Philip exported the PDF from the non-'_f2' deck, so the '_f2'
  # files parked in Morgue\ are dead and the Korean slides now come from that PDF.
  '10_Ratio_Analysis (Ch 5 English) KAIST - Students.pdf'                               = 'week3\english\Wk3_L10_slides.pdf'
  '10_Ratio_Analysis (Ch_5 Korean) KAIST - Students deploy.pdf'                         = 'week3\korean\Wk3_L10_slides.pdf'
  '10_Analyzing and Interpreting Financial Statements (Ch 5 English) - Students.pdf'    = 'week3\english\Wk3_L10_notes.pdf'
  '10_Analyzing and Interpreting Financial Statements (Ch 5 Korean) - Students.pdf'     = 'week3\korean\Wk3_L10_notes.pdf'

  # ---------------- WEEK 4 ----------------
  # Sessions 11 and 12 (Aug 15, Sat) are the Liberation Day holiday -- no class -- so
  # Week 4 opens at Lecture 13. PDF only, STUDENT versions only; the .pptx / .docx
  # sources sit beside these in Dropbox and are deliberately NOT mapped.
  # Watch the double space in "Operating  Assets", and the Korean notes filename is
  # missing its closing parenthesis ("Ch 8 Korean- Students") -- both kept verbatim.
  '13_LongTerm_Operating_Assets (Ch 8 English) _KAIST_Students.pdf'                     = 'week4\english\Wk4_L13_slides.pdf'
  '13_LongTerm_Operating_Assets (Ch 8 Korean) _KAIST_students.pdf'                      = 'week4\korean\Wk4_L13_slides.pdf'
  '13_Long-term Operating  Assets (Ch 8 English) - Students.pdf'                        = 'week4\english\Wk4_L13_notes.pdf'
  '13_Long-term Operating  Assets (Ch 8 Korean- Students.pdf'                           = 'week4\korean\Wk4_L13_notes.pdf'

  # ---------------- IN-CLASS ACTIVITIES (weeks/inclass.html) ----------------
  # Korean only for now; add an 'inclass\english\...' line when an English version exists.
  # Destination lecture numbers follow the CURRENT schedule, source prefixes the original
  # one -- match on the CHAPTER (the '7_' file is Ch 6, which is now Lecture 6).
  # STUDENT copies only -- never map the '... InClass Professor Copy.docx' files.
  '5_InClass Questions Adjusting Accounts for Financial Statements (Ch 3 Korean).docx'  = 'inclass\korean\IC_L5_activities.docx'
  '7_Revenue and Account Receivables - InClass Problems.pdf'                            = 'inclass\korean\IC_L6_activities.pdf'

  # ---------------- HOMEWORK (weeks/hw.html, weeks/hw-solutions.html) ----------------
  # Sources live in the 'HW' subfolder. Philip renumbered these prefixes on 2026-08-10
  # (Ch 7: 8_ -> 9_, Ch 5: 9_ -> 10_, Ch 8: 10_ -> 11_), so most now line up with the
  # lecture number -- but NOT all of them. Ch 8 is prefixed 11_ while the schedule page
  # calls it session 13 (sessions 11/12 are the Aug 15 holiday). So the rule still stands:
  # match on the CHAPTER, never on the prefix.
  #     5_  = Ch 3 -> Lecture 5      7_  = Ch 6 -> Lecture 6
  #     9_  = Ch 7 -> Lecture 9      10_ = Ch 5 -> Lecture 10     11_ = Ch 8 -> Lecture 13
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
  'HW\9_HW Inventory (Engslish Ch 7) - Problems.docx'                                  = 'hw\english\HW_L9_problems.docx'
  'HW\9_HW Inventory (Korean Ch 7) - Problems.docx'                                    = 'hw\korean\HW_L9_problems.docx'
  'HW\9_HW Inventory (Engslish Ch 7) - Solutions.docx'                                 = 'hw\english\HW_L9_solutions.docx'
  'HW\9_HW Inventory (Korean Ch 7) - Solutions.docx'                                   = 'hw\korean\HW_L9_solutions.docx'
  'HW\10_HW Ratio Analysis  (English Ch 5) - Problems.docx'                            = 'hw\english\HW_L10_problems.docx'
  'HW\10_HW Ratio Analysis  (Korean Ch 5) - Problems.docx'                             = 'hw\korean\HW_L10_problems.docx'
  'HW\10_HW Ratio Analysis  (English Ch 5) - Solutions.docx'                           = 'hw\english\HW_L10_solutions.docx'
  'HW\10_HW Ratio Analysis  (Korean Ch 5) - Solutions.docx'                            = 'hw\korean\HW_L10_solutions.docx'
  # Lecture 13 (Ch 8) restored 2026-08-14 with Week 4. Its Dropbox prefix moved 11_ -> 13_
  # (so it now matches the lecture number) and Philip exported PDFs, which is what we post
  # -- the .docx sources sit beside them and are deliberately NOT mapped.
  'HW\13_HW Long-Lived  Assets (English Ch 8) - Problems.pdf'                          = 'hw\english\HW_L13_problems.pdf'
  'HW\13_HW Long-Lived  Assets (Korean Ch 8) - Problems.pdf'                           = 'hw\korean\HW_L13_problems.pdf'
  'HW\13_HW Long-Lived  Assets (English Ch 8) - Solutions.pdf'                         = 'hw\english\HW_L13_solutions.pdf'
  'HW\13_HW Long-Lived  Assets (Korean Ch 8) - Solutions.pdf'                          = 'hw\korean\HW_L13_solutions.pdf'

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
    $KO_PROF    = [string]::Concat([char]0xAD50, [char]0xC218, [char]0xC6A9)   # 교수용

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
        # ---- Professor (answer-key) copies, released once the lecture has been taught ----
        @{ Filter = '99_Valuation_Primer (English) Professor.docx'
           Dest = 'week2\english\Wk2_L8_notes_prof.docx' }
        @{ Filter = '99_Valuation_Primer (*) Professor.docx'; Exclude = 'English'
           Dest = 'week2\korean\Wk2_L8_notes_prof.docx' }
        @{ Filter = '99_Valuation_Primer_*.docx'; Include = $KO_PROF
           Dest = 'inclass\korean\IC_L8_practice_solutions.docx' }
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

# ---------------- LECTURE 10 IN-CLASS (Ch 5 ratio analysis worksheet) ----------------
# Sources live in the '1_Ratio Analysis' subfolder. Every file name there is entirely
# Hangul, so nothing can be written as a literal in this ANSI-read .ps1 -- match on the
# extension and separate student from professor with the 학생용 / 교수용 code-point marker
# (same technique as the 0_Valuation block above).
# The 'morgue' subfolder holds superseded copies and is skipped (no -Recurse).
$ratSrc = Join-Path $src '1_Ratio Analysis'
if (Test-Path -LiteralPath $ratSrc) {
    $KO_STUDENT = [string]::Concat([char]0xD559, [char]0xC0DD, [char]0xC6A9)   # 학생용
    $KO_PROF    = [string]::Concat([char]0xAD50, [char]0xC218, [char]0xC6A9)   # 교수용

    $ratJobs = @(
        # STUDENT worksheet. The professor answer key sits beside it with the same .pdf
        # shape, so the two are separated by the Hangul marker.
        @{ Filter = '*.pdf'; Include = $KO_STUDENT; Exclude = $KO_PROF
           Dest = 'inclass\korean\IC_L10_activities.pdf' }
        # PROFESSOR answer key -- released now that Lecture 10 has been taught, same as the
        # Lecture 8 practice solutions above.
        @{ Filter = '*.pdf'; Include = $KO_PROF; Exclude = $KO_STUDENT
           Dest = 'inclass\korean\IC_L10_solutions.pdf' }
    )
    foreach ($job in $ratJobs) {
        $m = @(Get-ChildItem -LiteralPath $ratSrc -Filter $job.Filter -File -ErrorAction SilentlyContinue)
        if ($job.Exclude) { $m = @($m | Where-Object { $_.Name -notlike ('*' + $job.Exclude + '*') }) }
        if ($job.Include) { $m = @($m | Where-Object { $_.Name -like     ('*' + $job.Include + '*') }) }
        if ($m.Count -eq 0) { $missing += ('1_Ratio Analysis: ' + $job.Dest); continue }
        if ($m.Count -gt 1) { Write-Host ("  WARN {0} matched {1} files; using first" -f $job.Dest, $m.Count) -ForegroundColor Yellow }
        $dstPath = Join-Path $repo $job.Dest
        $dstDir  = Split-Path -Parent $dstPath
        if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }
        Copy-Item -LiteralPath $m[0].FullName -Destination $dstPath -Force
        $copied++
        Write-Host ("  OK  {0}" -f $job.Dest)
    }
}

# ---------------- EXAM PAGE (pop quiz over Lecture Notes 1-9) ----------------
# Sources live in 'Exam\Exams' -- a SIBLING of the notes folder, like 0_Admin. Every file
# name there is Hangul, so match on the extension and split student from answer key with
# the 학생용 / 답지 code-point markers (same technique as the 0_Valuation block above).
# The 'morgue' subfolder holds retired exams from earlier years and is skipped (no -Recurse).
# Korean only -- no English version of the quiz exists.
$exSrc = Join-Path (Split-Path -Parent $src) 'Exam\Exams'
if (Test-Path -LiteralPath $exSrc) {
    $KO_STUDENT = [string]::Concat([char]0xD559, [char]0xC0DD, [char]0xC6A9)   # 학생용
    $KO_KEY     = [string]::Concat([char]0xB2F5, [char]0xC9C0)                # 답지

    $exJobs = @(
        @{ Filter = '*.pdf'; Include = $KO_STUDENT; Exclude = $KO_KEY
           Dest = 'exam\korean\EX_popquiz_student.pdf' }
        @{ Filter = '*.pdf'; Include = $KO_KEY; Exclude = $KO_STUDENT
           Dest = 'exam\korean\EX_popquiz_key.pdf' }
    )
    foreach ($job in $exJobs) {
        $m = @(Get-ChildItem -LiteralPath $exSrc -Filter $job.Filter -File -ErrorAction SilentlyContinue)
        if ($job.Exclude) { $m = @($m | Where-Object { $_.Name -notlike ('*' + $job.Exclude + '*') }) }
        if ($job.Include) { $m = @($m | Where-Object { $_.Name -like     ('*' + $job.Include + '*') }) }
        if ($m.Count -eq 0) { $missing += ('Exam\Exams: ' + $job.Dest); continue }
        if ($m.Count -gt 1) { Write-Host ("  WARN {0} matched {1} files; using first" -f $job.Dest, $m.Count) -ForegroundColor Yellow }
        $dstPath = Join-Path $repo $job.Dest
        $dstDir  = Split-Path -Parent $dstPath
        if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }
        Copy-Item -LiteralPath $m[0].FullName -Destination $dstPath -Force
        $copied++
        Write-Host ("  OK  {0}" -f $job.Dest)
    }
}

# ---------------- WEEK 0 (Course Resources: syllabus + accounting terminology) ----------------
# The terminology glossaries live in the "Other Resources web" subfolder, split into
# English / 한국어 language subfolders. The Korean folder name contains Hangul, so we locate
# it at runtime to keep this ANSI-read script free of non-ASCII literals (same reasoning as
# the video block above).
#
# The SYLLABUS is different: as of 2026-08-14 it comes straight from '0_Admin' (a SIBLING
# of the notes folder, not a subfolder), which is where Philip actually edits it -- the
# copies under 'Other Resources web' were a stale second home and that copy step is gone.
# It publishes as PDF now, like the lectures. Both jobs take the MOST RECENTLY MODIFIED
# match rather than a dated file name, so a re-revision ("... revised 8_14_2026.pdf" ->
# some later name) is picked up without editing this script.
$w0src = Join-Path $src 'Other Resources web'
$admin = Join-Path (Split-Path -Parent $src) '0_Admin'
if (Test-Path -LiteralPath $w0src) {
    $w0en    = Join-Path $w0src 'English'
    $w0koDir = Get-ChildItem -LiteralPath $w0src -Directory | Where-Object { $_.Name -ne 'English' } | Select-Object -First 1
    $w0ko    = if ($w0koDir) { $w0koDir.FullName } else { $null }
    $w0admin = if (Test-Path -LiteralPath $admin) { $admin } else { $null }

    $w0jobs = @(
        # Syllabus -- English name is ASCII; the Korean one is '..._강의계획서 한국어 ....pdf',
        # so it is matched as "the BAF50002 PDF that is not the English one".
        @{ Dir = $w0admin; Filter = 'BAF50002_Syllabus*.pdf'; Newest = $true
           Dest = 'week0\english\Wk0_Syllabus.pdf' }
        @{ Dir = $w0admin; Filter = 'BAF50002_*.pdf'; Exclude = 'Syllabus'; Newest = $true
           Dest = 'week0\korean\Wk0_Syllabus.pdf' }
        @{ Dir = $w0en; Filter = '*Terminology*English*.xlsx'; Dest = 'week0\english\Wk0_Terminology.xlsx' }
        @{ Dir = $w0ko; Filter = '*Terminology*Korean*.xlsx';  Dest = 'week0\korean\Wk0_Terminology.xlsx' }
    )
    foreach ($job in $w0jobs) {
        if (-not $job.Dir) { $missing += ('week0: ' + $job.Filter); continue }
        $m = @(Get-ChildItem -LiteralPath $job.Dir -Filter $job.Filter -File -ErrorAction SilentlyContinue)
        if ($job.Exclude) { $m = @($m | Where-Object { $_.Name -notlike ('*' + $job.Exclude + '*') }) }
        if ($m.Count -eq 0) { $missing += ('week0: ' + $job.Filter); continue }
        if ($job.Newest) {
            $m = @($m | Sort-Object LastWriteTime -Descending)
        } elseif ($m.Count -gt 1) {
            Write-Host ("  WARN {0} matched {1} files; using first" -f $job.Filter, $m.Count) -ForegroundColor Yellow
        }
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

# Word/PowerPoint sometimes stamp "Hong, Philip Keejae" into an exported PDF's /Author and
# XMP metadata, and copying from Dropbox brings that stamp along -- so scrub the published
# copies on every sync. Idempotent and length-preserving; leaves the Dropbox originals alone.
$pdfScrub = Join-Path $PSScriptRoot 'scripts\strip_pdf_metadata.ps1'
if (Test-Path -LiteralPath $pdfScrub) {
    Write-Host ""
    Write-Host "Scrubbing PDF metadata..."
    & $pdfScrub -Path $repo -Recurse
}

if ($missing.Count -gt 0) {
    Write-Host ""
    Write-Host "MISSING SOURCE FILES (not found in Dropbox folder):" -ForegroundColor Yellow
    $missing | Sort-Object -Unique | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Yellow }
}
