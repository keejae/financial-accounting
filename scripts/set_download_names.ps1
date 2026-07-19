# Sets the download="" (student-facing filename) on each week page to the
# original descriptive source filename, while leaving href paths untouched.
$repo = "C:\Users\hong1p\Documents\GitHub\financial-accounting"

# hrefPath (as written in HTML)  =>  original source filename (raw)
$map = [ordered]@{
  # ---- WEEK 1 ----
  "../files/week1/english/Wk1_L1_slides_pt1.pptx" = "1_Introduction to Accounting - Color Accounting (Korean only) - Students deploy.pptx"
  "../files/week1/korean/Wk1_L1_slides_pt1.pptx"  = "1_Introduction to Accounting - Color Accounting (Korean only) - Students deploy.pptx"
  "../files/week1/english/Wk1_L1_slides_pt2.pptx" = "2_1_Classic Transactions Color Accounting KAIST( Korean only)  Students deploy.pptx"
  "../files/week1/korean/Wk1_L1_slides_pt2.pptx"  = "2_1_Classic Transactions Color Accounting KAIST( Korean only)  Students deploy.pptx"
  "../files/week1/english/Wk1_L1_notes.docx" = "1_to_2_Activity Book Color Accounting (Ch 0  English) - Students.docx"
  "../files/week1/korean/Wk1_L1_notes.docx"  = "1_to_2_Activity Book Color Accounting (Ch 0 Korean) - Students.docx"
  "../files/week1/english/Wk1_L2_slides.pptx" = "2_2_Journals and T-Accounts KAIST (Korean only) Students deploy.pptx"
  "../files/week1/korean/Wk1_L2_slides.pptx"  = "2_2_Journals and T-Accounts KAIST (Korean only) Students deploy.pptx"
  "../files/week1/english/Wk1_L3_slides.pptx" = "3_Introducting Financial Accounting  (Ch 1 English) KAIST - Professor.pptx"
  "../files/week1/korean/Wk1_L3_slides.pptx"  = "3_Introducing_Financial_Accounting (Ch 1 Korean) KAIST - Students deploy.pptx"
  "../files/week1/english/Wk1_L3_notes.docx" = "3_Introducing Financial Accounting  (Ch 1  English)- Students.docx"
  "../files/week1/korean/Wk1_L3_notes.docx"  = "3_Introducing Financial Accounting  (Ch 1  Korean)- Students.docx"
  "../files/week1/english/Wk1_L4_slides.pptx" = "4_Constructing Financial Statements (Ch 2 English) KAIST- Professor.pptx"
  "../files/week1/korean/Wk1_L4_slides.pptx"  = "4_Constructing_Financial_Statements (Ch 2 Korean) KAIST- Students deploy.pptx"
  "../files/week1/english/Wk1_L4_notes.docx" = "4_Constructing Financial Statements (Ch 2 English) - Students.docx"
  "../files/week1/korean/Wk1_L4_notes.docx"  = "4_Constructing Financial Statements (Ch 2 Korean) - Students.docx"
  # ---- WEEK 2 ----
  "../files/week2/english/Wk2_L5_slides.pptx" = "5_Adjusting Accounts for Financial Statements (Ch 3 English) KAIST - Professor.pptx"
  "../files/week2/korean/Wk2_L5_slides.pptx"  = "5_Adjusting_Accounts_for_Financial_Statements (Ch_3_Korean)  KAIST - Students deploy.pptx"
  "../files/week2/english/Wk2_L5_notes.docx" = "5_Adjusting Accounts for Financial Statements (Ch 3 English) - Students.docx"
  "../files/week2/korean/Wk2_L5_notes.docx"  = "5_Adjusting Accounts for Financial Statements (Ch 3 Korean) - Students.docx"
  "../files/week2/english/Wk2_L6_slides.pptx" = "6_Statement of Cash Flows (Ch 4 English) KAIST - Professor.pptx"
  "../files/week2/korean/Wk2_L6_slides.pptx"  = "6_Statement of Cash Flows (Ch 4 Korean) KAIST - Students deploy.pptx"
  "../files/week2/english/Wk2_L6_notes.docx" = "6_Statement of Cash Flows short (Ch 4 English) -  Students.docx"
  "../files/week2/korean/Wk2_L6_notes.docx"  = "6_Statement of Cash Flows short (Ch 4 Korean) - Students.docx"
  "../files/week2/english/Wk2_L7_slides.pptx" = "7_Revenue_and_Receivables (Ch 6 English) KAIST Professor.pptx"
  "../files/week2/korean/Wk2_L7_slides.pptx"  = "7_Revenue_and_Receivables (Ch 6 Korean) KAIST Students deploy.pptx"
  "../files/week2/english/Wk2_L7_notes.docx" = "7_Revenue and Account Receivables (Ch 6 English)  shorter version - Students.docx"
  "../files/week2/korean/Wk2_L7_notes.docx"  = "7_Revenue and Account Receivables (Ch 6 Korean) shorter version - Students.docx"
  "../files/week2/english/Wk2_L8_slides.pptx" = "8_Inventory (Ch 7 English) KAIS ProfessorT.pptx"
  "../files/week2/korean/Wk2_L8_slides.pptx"  = "8_Inventory (Ch 7 Korean) KAIST -Students deploy.pptx"
  "../files/week2/english/Wk2_L8_notes.docx" = "8_Reporting and Analyzing Inventory (Ch 7 English) - Students.docx"
  "../files/week2/korean/Wk2_L8_notes.docx"  = "8_Reporting and Analyzing Inventory (Ch 7 Korean) - Students.docx"
  # ---- WEEK 3 ----
  "../files/week3/english/Wk3_L9_slides.pptx" = "9_Ratio_Analysis (Ch 5 English) KAIST - Professor.pptx"
  "../files/week3/korean/Wk3_L9_slides.pptx"  = "9_Ratio_Analysis (Ch_5 Korean) KAIST - Students deploy.pptx"
  "../files/week3/english/Wk3_L9_notes.docx" = "9_Analyzing and Interpreting Financial Statements (Ch 5 English) - Students.docx"
  "../files/week3/korean/Wk3_L9_notes.docx"  = "9_Analyzing and Interpreting Financial Statements (Ch 5 Korean) - Students.docx"
}

function Clean-Name([string]$n){
  $n = $n -replace '_', ' '                    # underscores -> spaces (readability)
  $n = $n -replace '\s+', ' '                  # collapse multiple spaces
  $n = $n -replace 'Introducting', 'Introducing'
  $n = $n -replace 'KAIS ProfessorT', 'KAIST Professor'
  return $n.Trim()
}

$pages = @{
  "week1" = Join-Path $repo "weeks\week1.html"
  "week2" = Join-Path $repo "weeks\week2.html"
  "week3" = Join-Path $repo "weeks\week3.html"
}
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$content = @{}
foreach($k in $pages.Keys){ $content[$k] = [System.IO.File]::ReadAllText($pages[$k], [System.Text.Encoding]::UTF8) }

$applied = 0
foreach($entry in $map.GetEnumerator()){
  $href = $entry.Key
  $dl   = Clean-Name $entry.Value
  $wk = if($href -match '/week1/'){"week1"} elseif($href -match '/week2/'){"week2"} else {"week3"}
  $escHref = [regex]::Escape($href)
  $pattern = '(href="' + $escHref + '" download=")[^"]*(")'
  $newContent = [regex]::Replace($content[$wk], $pattern, { param($m) $m.Groups[1].Value + $dl + $m.Groups[2].Value })
  if($newContent -ne $content[$wk]){ $applied++ } else { Write-Host "NO MATCH: $href" -ForegroundColor Yellow }
  $content[$wk] = $newContent
}

foreach($k in $pages.Keys){ [System.IO.File]::WriteAllText($pages[$k], $content[$k], $utf8NoBom) }
Write-Host "Applied $applied download-name updates."
