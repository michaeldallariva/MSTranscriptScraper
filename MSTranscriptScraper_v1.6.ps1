############################################################################################################################################################################################################################################################
# Microsoft Certification Transcript Report Generator 
# GitHub link : https://github.com/michaeldallariva
# Version : v1.6
# Author : Michael DALLA RIVA, with the help of some AI
# Date : 01 Mar 2025
#
# Purpose:
# This script reads a saved transcript text file and generates a nicely formatted HTML report
# 
# Go to your Microsoft Education workdspace, share your transcript. Go to the shared link. Save the page as a .TXT file on your desktop. Run this script from your desktop (Or any other dedicated folder)
# 
# License :
# Feel free to use for any purpose, personal or commercial.
#
############################################################################################################################################################################################################################################################

function Clean-CertificationTitle {
    param (
        [string]$title
    )
    
    $title = $title -replace [regex]::Escape([char]0xE2) + "[\u0084][\u0162]", "(TM)"
    
    $title = $title -replace [char]0xC2 + [char]0xAE, "(R)"
    
    $title = $title -replace [char]0xC2 + [char]0xA9, "(C)"
    
    $title = [System.Text.RegularExpressions.Regex]::Replace($title, "[^\x20-\x7E]", "")
    
    $title = $title -replace "\s+", " "
    
    return $title.Trim()
}

function Extract-HistoricalCertifications {
    param (
        [string]$textContent
    )
    
    $historicalCerts = @()
    
    if ($textContent -match "(?s)(Historical certifications.*?)Show less") {
        $histSection = $matches[1]
    } elseif ($textContent -match "(?s)(Historical certifications.*)$") {
        $histSection = $matches[1]
    } else {
        Write-Host "Historical certifications section not found." -ForegroundColor Yellow
        return $historicalCerts
    }
    
    $blocks = $histSection -split "(?m)^\s*\d+\.\s*$"
    
    foreach ($block in $blocks) {
        if ([string]::IsNullOrWhiteSpace($block) -or $block.Trim() -eq "Historical certifications") {
            continue
        }
        
        $titleMatch = [regex]::Match($block, "(?s)Certification title\s+(.+?)(?=\s+Earned on|\s+Certification number)")
        if ($titleMatch.Success) {
            $title = $titleMatch.Groups[1].Value.Trim()
            
            $earnedOnMatch = [regex]::Match($block, "Earned on\s+([^\s]+\s+[^\s]+\s+[^\s]+)")
            $earnedOn = if ($earnedOnMatch.Success) { $earnedOnMatch.Groups[1].Value.Trim() } else { "" }
            
            $expiredOnMatch = [regex]::Match($block, "Expired on\s+([^\s]+\s+[^\s]+\s+[^\s]+)")
            $expiredOn = if ($expiredOnMatch.Success) { $expiredOnMatch.Groups[1].Value.Trim() } else { "" }
            
            $certMatch = [regex]::Match($block, "(?s)Certification number\s+(.+?)(?=\s+State)")
            $certNumber = if ($certMatch.Success) { $certMatch.Groups[1].Value.Trim() } else { "" }
            
            $stateMatch = [regex]::Match($block, "(?s)State\s+(.+?)(\r?\n|$)")
            $state = if ($stateMatch.Success) { $stateMatch.Groups[1].Value.Trim() } else { "" }
            
            if ($state -match "^(Expired|Retired)") {
                $state = $matches[1]
            }
            
            if ($title -and $certNumber -and $state) {
                $historicalCerts += [PSCustomObject]@{
                    Title = $title
                    CertificationNumber = $certNumber
                    EarnedOn = $earnedOn
                    ExpiredOn = $expiredOn
                    State = $state
                }
            }
        }
    }
    
    if ($historicalCerts.Count -eq 0 -or ($historicalCerts | Where-Object { $_.EarnedOn -eq "" }).Count -gt 0) {
        Write-Host "Trying to find missing earned dates..." -ForegroundColor Yellow
        
        $earnedDatePattern = "(?m)Earned on\s+(\d+\s+[A-Za-z]+\s+\d+)(?:\s+Expired on\s+(\d+\s+[A-Za-z]+\s+\d+))?"
        $earnedDateMatches = [regex]::Matches($histSection, $earnedDatePattern)
        
        foreach ($match in $earnedDateMatches) {
            $earnedOn = $match.Groups[1].Value.Trim()
            $expiredOn = if ($match.Groups[2].Success) { $match.Groups[2].Value.Trim() } else { "" }
            
            $context = $histSection.Substring(0, $match.Index + $match.Length + 100) # Look at some context after the match
            $certNumberMatch = [regex]::Match($context, "([A-Z0-9]+-[A-Z0-9]+)")
            $titleMatch = [regex]::Match($context.Substring(0, $match.Index), "(?s)Certification title\s+(.+?)(?=\s+Earned on)")
            
            if ($certNumberMatch.Success) {
                $certNumber = $certNumberMatch.Groups[1].Value.Trim()
                
                $existingCert = $historicalCerts | Where-Object { $_.CertificationNumber -eq $certNumber -and $_.EarnedOn -eq "" }
                
                if ($existingCert) {
                    $existingCert.EarnedOn = $earnedOn
                    $existingCert.ExpiredOn = $expiredOn
                }
            }
        }
    }
    
    $missingEarnedDates = $historicalCerts | Where-Object { $_.EarnedOn -eq "" }
    if ($missingEarnedDates.Count -gt 0) {
        Write-Host "Searching for more earned dates using different pattern..." -ForegroundColor Yellow
        
        foreach ($cert in $missingEarnedDates) {
            $certPattern = "(?s).{0,500}$($cert.CertificationNumber).{0,500}"
            $certContext = [regex]::Match($histSection, $certPattern).Value
            
            $earnedOnMatch = [regex]::Match($certContext, "Earned on\s+(\d+\s+[A-Za-z]+\s+\d+)")
            if ($earnedOnMatch.Success) {
                $cert.EarnedOn = $earnedOnMatch.Groups[1].Value.Trim()
            }
        }
    }
    
    return $historicalCerts
}
function Extract-MCTData {
    param (
        [string]$relevantText
    )
    
    $mctHistory = @()
    
    if ($relevantText -like "*Microsoft Certified Trainer history*") {
        Write-Host "Using direct extraction for MCT history..." -ForegroundColor Yellow
        
        if ($relevantText -match "(?s)(Microsoft Certified Trainer history.+?Historical certifications)") {
            $mctBlock = $matches[1]
            
            $dateMatches = [regex]::Matches($mctBlock, "(\d+\s+[A-Za-z]+\s+\d+)")
            
            if ($dateMatches.Count -ge 2) {
                $activeFrom = $dateMatches[0].Value.Trim()
                $to = $dateMatches[1].Value.Trim()
                
                $mctHistory += [PSCustomObject]@{
                    Title = "MCT History"
                    ActiveFrom = $activeFrom
                    To = $to
                }
                
                Write-Host "Successfully extracted MCT history directly." -ForegroundColor Green
            }
        }
    }
    
    return $mctHistory
}

function Get-StyleSheet {
    return @"
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            color: #333;
            line-height: 1.6;
            background-color: #f9f9f9;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background-color: #fff;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            border-radius: 5px;
        }
        h1, h2, h3 {
            color: #0078d4;
            margin-top: 0;
        }
        h1 {
            font-size: 28px;
            border-bottom: 2px solid #0078d4;
            padding-bottom: 10px;
        }
        h2 {
            font-size: 22px;
            margin-top: 30px;
            padding-bottom: 5px;
            border-bottom: 1px solid #eaeaea;
        }
        .header {
            background-color: #0078d4;
            color: white;
            padding: 15px;
            border-radius: 5px 5px 0 0;
            margin-bottom: 20px;
        }
        .header h1 {
            color: white;
            border-bottom: none;
            margin: 0;
            padding: 0;
        }
        .info-section {
            margin-bottom: 30px;
            background-color: white;
            padding: 20px;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        }
        .info-item {
            margin-bottom: 10px;
        }
        .info-label {
            font-weight: bold;
            display: inline-block;
            width: 200px;
            color: #555;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }
        th {
            background-color: #0078d4;
            color: white;
            text-align: left;
            padding: 12px;
        }
        td {
            padding: 10px;
            border: 1px solid #ddd;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        .footer {
            margin-top: 30px;
            text-align: center;
            font-size: 12px;
            color: #666;
            padding: 10px;
            border-top: 1px solid #eaeaea;
        }
        .stat-box {
            background-color: #f0f8ff;
            border-left: 4px solid #0078d4;
            padding: 10px 15px;
            margin-bottom: 15px;
            border-radius: 0 3px 3px 0;
        }
        .stat-number {
            font-size: 24px;
            font-weight: bold;
            color: #0078d4;
        }
        .stat-label {
            font-size: 14px;
            color: #555;
        }
        .stats-container {
            display: flex;
            justify-content: space-between;
            margin-bottom: 20px;
        }
        .stat-box {
            flex: 1;
            margin-right: 10px;
        }
        .stat-box:last-child {
            margin-right: 0;
        }
        @media print {
            body {
                padding: 0;
                background-color: white;
            }
            .container {
                box-shadow: none;
            }
        }
    </style>
"@
}

function Get-HtmlHeader {
    param (
        [string]$Title,
        [string]$Name
    )
    
    return @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$Title</title>
    $(Get-StyleSheet)
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Microsoft Certification Transcript</h1>
        </div>
"@
}

function Get-HtmlFooter {
    $date = Get-Date -Format "yyyy-MM-dd"
    return @"
        <div class="footer">
            <p>Report generated on $date</p>
        </div>
    </div>
</body>
</html>
"@
}

function Extract-Certifications {
    param (
        [string]$textContent,
        [string]$sectionType # "active" or "historical"
    )
    
    $certifications = @()
    
    if ($sectionType -eq "historical") {
        $histPattern = @"
(?s)Microsoft Certified:.*?
Earned on\s+(\d+\s+[A-Za-z]+\s+\d+)\s+Expired on\s+(\d+\s+[A-Za-z]+\s+\d+).*?
Certification number.*?
([A-Z0-9]+-[A-Z0-9]+).*?
State.*?
(Expired|Retired)
"@
        
        $histMatches = [regex]::Matches($textContent, $histPattern)
        
        foreach ($match in $histMatches) {
            $beforeEarnedOn = $textContent.Substring(0, $match.Index + $match.Groups[0].Index)
            $lines = $beforeEarnedOn -split "\r?\n"
            $titleLine = $lines | Where-Object { $_ -match "Microsoft Certified" } | Select-Object -Last 1
            
            if ($titleLine) {
                $title = $titleLine.Trim()
                $earnedOn = $match.Groups[1].Value.Trim()
                $expiredOn = $match.Groups[2].Value.Trim()
                $certNumber = $match.Groups[3].Value.Trim()
                $state = $match.Groups[4].Value.Trim()
                
                $certifications += [PSCustomObject]@{
                    Title = $title
                    CertificationNumber = $certNumber
                    EarnedOn = $earnedOn
                    ExpiredOn = $expiredOn
                    State = $state
                }
            }
        }
        
        if ($certifications.Count -eq 0) {
            $blocks = $textContent -split "\r?\n\r?\n"
            
            foreach ($block in $blocks) {
                if ($block -match "Microsoft Certified" -or $block -match "Microsoft®") {
                    $titleMatch = [regex]::Match($block, "(Microsoft[^$\r\n]+)")
                    $earnedOnMatch = [regex]::Match($block, "Earned on\s+([^$\r\n]+)")
                    $expiredOnMatch = [regex]::Match($block, "Expired on\s+([^$\r\n]+)")
                    $certNumberMatch = [regex]::Match($block, "([A-Z0-9]+-[A-Z0-9]+)")
                    $stateMatch = [regex]::Match($block, "(Expired|Retired)")
                    
                    if ($titleMatch.Success -and $certNumberMatch.Success) {
                        $title = $titleMatch.Groups[1].Value.Trim()
                        $earnedOn = if ($earnedOnMatch.Success) { $earnedOnMatch.Groups[1].Value.Trim() } else { "" }
                        $expiredOn = if ($expiredOnMatch.Success) { $expiredOnMatch.Groups[1].Value.Trim() } else { "" }
                        $certNumber = $certNumberMatch.Groups[1].Value.Trim()
                        $state = if ($stateMatch.Success) { $stateMatch.Groups[1].Value.Trim() } else { "Retired" }
                        
                        $certifications += [PSCustomObject]@{
                            Title = $title
                            CertificationNumber = $certNumber
                            EarnedOn = $earnedOn
                            ExpiredOn = $expiredOn
                            State = $state
                        }
                    }
                }
            }
        }
        
        if ($certifications.Count -eq 0) {
            $tablePattern = @"
(?m)^(Microsoft[^\r\n]+?)
(?:\s+Earned on\s+([^\r\n]+?)(?:\s+Expired on\s+([^\r\n]+?))?)?
\s+([A-Z0-9]+-[A-Z0-9]+)\s+(Expired|Retired)$
"@
            
            $tableMatches = [regex]::Matches($textContent, $tablePattern)
            
            foreach ($match in $tableMatches) {
                $title = $match.Groups[1].Value.Trim()
                $earnedOn = $match.Groups[2].Value.Trim()
                $expiredOn = $match.Groups[3].Value.Trim()
                $certNumber = $match.Groups[4].Value.Trim()
                $state = $match.Groups[5].Value.Trim()
                
                $certifications += [PSCustomObject]@{
                    Title = $title
                    CertificationNumber = $certNumber
                    EarnedOn = $earnedOn
                    ExpiredOn = $expiredOn
                    State = $state
                }
            }
        }
        
        if ($certifications.Count -eq 0) {
            $lines = $textContent -split "\r?\n"
            
            for ($i = 0; $i -lt $lines.Count; $i++) {
                $line = $lines[$i].Trim()
                
                if ($line -match "^Microsoft") {
                    $title = $line
                    $earnedOn = ""
                    $expiredOn = ""
                    $certNumber = ""
                    $state = ""
                    
                    for ($j = $i + 1; $j -lt [Math]::Min($i + 10, $lines.Count); $j++) {
                        $nextLine = $lines[$j].Trim()
                        
                        if ($nextLine -match "Earned on\s+(.+?)(?:\s+Expired on\s+(.+?))?$") {
                            $earnedOn = $matches[1].Trim()
                            $expiredOn = if ($matches[2]) { $matches[2].Trim() } else { "" }
                        }
                        elseif ($nextLine -match "([A-Z0-9]+-[A-Z0-9]+)\s+(Expired|Retired)$") {
                            $certNumber = $matches[1].Trim()
                            $state = $matches[2].Trim()
                        }
                    }
                    
                    if ($title -and $certNumber -and $state) {
                        $certifications += [PSCustomObject]@{
                            Title = $title
                            CertificationNumber = $certNumber
                            EarnedOn = $earnedOn
                            ExpiredOn = $expiredOn
                            State = $state
                        }
                    }
                }
            }
        }
        
        return $certifications
    }
    
    $pattern = switch ($sectionType) {
        "active" {
            "(?m)^\s*Certification title\s*$\s*^(.+?)$\s*^\s*Certification number\s*$\s*^(.+?)$\s*^\s*Earned on\s*$\s*^(.+?)$\s*^\s*Expires on\s*$\s*^(.+?)$"
        }
        "historical" {
            "(?m)^\s*Certification title\s*$\s*^(.+?)$\s*^Earned on (.+?)(Expired on (.+?))?$\s*^\s*Certification number\s*$\s*^(.+?)$\s*^\s*State\s*$\s*^(.+?)$"
        }
    }
    
    $tablePattern = switch ($sectionType) {
        "active" {
            "(?m)^(.+?)\s{2,}([A-Z0-9]+-[A-Z0-9]+)\s{2,}(\d+\s+[A-Za-z]+\s+\d+)\s{2,}(N/A)\s*$"
        }
        "historical" {
            "(?m)^(.+?)\s{2,}([A-Z0-9]+-[A-Z0-9]+)\s{2,}(Expired|Retired)\s*$"
        }
    }
    
    $matches = [regex]::Matches($textContent, $pattern)
    
    if ($matches.Count -gt 0) {
        foreach ($match in $matches) {
            if ($sectionType -eq "active") {
                $title = $match.Groups[1].Value.Trim()
                $certNumber = $match.Groups[2].Value.Trim()
                $earnedOn = $match.Groups[3].Value.Trim()
                $expiresOn = $match.Groups[4].Value.Trim()
                
                $certifications += [PSCustomObject]@{
                    Title = $title
                    CertificationNumber = $certNumber
                    EarnedOn = $earnedOn
                    ExpiresOn = $expiresOn
                }
            }
            else {
                $title = $match.Groups[1].Value.Trim()
                $earnedOn = $match.Groups[2].Value.Trim()
                $expiredOn = if ($match.Groups[4].Success) { $match.Groups[4].Value.Trim() } else { "" }
                $certNumber = $match.Groups[5].Value.Trim()
                $state = $match.Groups[6].Value.Trim()
                
                $certifications += [PSCustomObject]@{
                    Title = $title
                    CertificationNumber = $certNumber
                    EarnedOn = $earnedOn
                    ExpiredOn = $expiredOn
                    State = $state
                }
            }
        }
    }
    
    if ($certifications.Count -eq 0) {
        $tableMatches = [regex]::Matches($textContent, $tablePattern)
        
        foreach ($match in $tableMatches) {
            if ($sectionType -eq "active") {
                $title = $match.Groups[1].Value.Trim()
                $certNumber = $match.Groups[2].Value.Trim()
                $earnedOn = $match.Groups[3].Value.Trim()
                $expiresOn = $match.Groups[4].Value.Trim()
                
                $certifications += [PSCustomObject]@{
                    Title = $title
                    CertificationNumber = $certNumber
                    EarnedOn = $earnedOn
                    ExpiresOn = $expiresOn
                }
            }
            else {
                $title = $match.Groups[1].Value.Trim()
                $certNumber = $match.Groups[2].Value.Trim()
                $state = $match.Groups[3].Value.Trim()
                
                $datesMatch = [regex]::Match($title, "Earned on\s+(.+?)\s+Expired on\s+(.+)")
                $earnedOn = if ($datesMatch.Success) { $datesMatch.Groups[1].Value.Trim() } else { "" }
                $expiredOn = if ($datesMatch.Success) { $datesMatch.Groups[2].Value.Trim() } else { "" }
                
                 $title = $title -replace "Earned on.+?$", ""
                
                $certifications += [PSCustomObject]@{
                    Title = $title.Trim()
                    CertificationNumber = $certNumber
                    EarnedOn = $earnedOn
                    ExpiredOn = $expiredOn
                    State = $state
                }
            }
        }
    }
    
    if (($sectionType -eq "historical") -and ($certifications.Count -eq 0)) {
        $histCertPattern = [regex]::new(@"
(?m)Certification title\s*$\s*^(.+?)$\s*^\s*Earned on\s+(.+?)$\s*^\s*Expired on\s+(.+?)$\s*^\s*Certification number\s*$\s*^(.+?)$\s*^\s*State\s*$\s*^(.+?)$
"@, [System.Text.RegularExpressions.RegexOptions]::Singleline)
        
        $histCertMatches = $histCertPattern.Matches($textContent)
        
        foreach ($match in $histCertMatches) {
            $title = $match.Groups[1].Value.Trim()
            $earnedOn = $match.Groups[2].Value.Trim()
            $expiredOn = $match.Groups[3].Value.Trim()
            $certNumber = $match.Groups[4].Value.Trim()
            $state = $match.Groups[5].Value.Trim()
            
            $certifications += [PSCustomObject]@{
                Title = $title
                CertificationNumber = $certNumber
                EarnedOn = $earnedOn
                ExpiredOn = $expiredOn
                State = $state
            }
        }
        
        if ($certifications.Count -eq 0) {
            $blocks = $textContent -split "(?:\r?\n){2,}"
            
            for ($i = 0; $i -lt $blocks.Count; $i++) {
                $block = $blocks[$i]
                
                if ($block -notmatch "Certification title|[A-Z0-9]+-[A-Z0-9]+") {
                    continue
                }
                
                $titleMatch = [regex]::Match($block, "(?m)^(.+?)$")
                if ($titleMatch.Success) {
                    $title = $titleMatch.Groups[1].Value.Trim()
                    
                    $certNumber = ""
                    $state = ""
                    $earnedOn = ""
                    $expiredOn = ""
                    
                    $datesMatch = [regex]::Match($block, "Earned on\s+(.+?)(?:\s+Expired on\s+(.+))?$")
                    if ($datesMatch.Success) {
                        $earnedOn = $datesMatch.Groups[1].Value.Trim()
                        $expiredOn = if ($datesMatch.Groups[2].Success) { $datesMatch.Groups[2].Value.Trim() } else { "" }
                        
                        $title = $title -replace "Earned on.+$", ""
                    }
                    
                    for ($j = $i + 1; $j -lt [Math]::Min($i + 5, $blocks.Count); $j++) {
                        $nextBlock = $blocks[$j]
                        
                        $certMatch = [regex]::Match($nextBlock, "([A-Z0-9]+-[A-Z0-9]+)")
                        if ($certMatch.Success) {
                            $certNumber = $certMatch.Groups[1].Value.Trim()
                        }
                        
                        $stateMatch = [regex]::Match($nextBlock, "(Expired|Retired)")
                        if ($stateMatch.Success) {
                            $state = $stateMatch.Groups[1].Value.Trim()
                        }
                    }
                    
                    if ($title -and $certNumber -and $state) {
                        $certifications += [PSCustomObject]@{
                            Title = $title.Trim()
                            CertificationNumber = $certNumber
                            EarnedOn = $earnedOn
                            ExpiredOn = $expiredOn
                            State = $state
                        }
                    }
                }
            }
        }
        
        if ($certifications.Count -eq 0) {
            $historicalBlockPattern = [regex]::new(@"
([\w\s:,]+?)

Earned on\s+(.+?)(?:\s+Expired on\s+(.+?))?

\s+([A-Z0-9]+-[A-Z0-9]+)\s+(Expired|Retired)
"@, [System.Text.RegularExpressions.RegexOptions]::Singleline)
            
            $historicalBlockMatches = $historicalBlockPattern.Matches($textContent)
            
            foreach ($match in $historicalBlockMatches) {
                $title = $match.Groups[1].Value.Trim()
                $earnedOn = $match.Groups[2].Value.Trim()
                $expiredOn = if ($match.Groups[3].Success) { $match.Groups[3].Value.Trim() } else { "" }
                $certNumber = $match.Groups[4].Value.Trim()
                $state = $match.Groups[5].Value.Trim()
                
                $title = $title -replace "^(Historical certifications|Certification title|Certification number|State)\s+", ""
                
                $certifications += [PSCustomObject]@{
                    Title = $title
                    CertificationNumber = $certNumber
                    EarnedOn = $earnedOn
                    ExpiredOn = $expiredOn
                    State = $state
                }
            }
        }
    }
    
    return $certifications
}

function Extract-Exams {
    param (
        [string]$textContent
    )
    
    $exams = @()
    
    $numberedPattern = "(?m)(?:\s*(\d+)\.\s*)?(?:\s*Exam title\s*\r?\n\s*((?:.+\r?\n?)+?)(?=\s*Exam number))\s*Exam number\s*\r?\n\s*([A-Z0-9-]+|[0-9]{2,3})\s*\r?\n\s*Passed date\s*\r?\n\s*(\d+\s+[A-Za-z]+\s+\d+)"
    $numberedMatches = [regex]::Matches($textContent, $numberedPattern)
    
    foreach ($match in $numberedMatches) {
        $rawTitle = $match.Groups[2].Value
        $cleanTitle = ($rawTitle -replace "\r?\n", " ").Trim()
        
        $examNumber = $match.Groups[3].Value.Trim()
        $passedDate = $match.Groups[4].Value.Trim()
        
        if ($cleanTitle -and $examNumber -and $passedDate) {
            $exists = $exams | Where-Object { $_.ExamNumber -eq $examNumber }
            
            if (-not $exists) {
                $exams += [PSCustomObject]@{
                    Title = $cleanTitle
                    ExamNumber = $examNumber
                    PassedDate = $passedDate
                }
            }
        }
    }
    
    if ($exams.Count -lt 5) {
        $pattern = "(?m)^\s*Exam title\s*$\s*^(.+?)$\s*^\s*Exam number\s*$\s*^(.+?)$\s*^\s*Passed date\s*$\s*^(.+?)$"
        $matches = [regex]::Matches($textContent, $pattern)
        
        foreach ($match in $matches) {
            $title = $match.Groups[1].Value.Trim()
            $examNumber = $match.Groups[2].Value.Trim()
            $passedDate = $match.Groups[3].Value.Trim()
            
            $exists = $exams | Where-Object { $_.ExamNumber -eq $examNumber }
            
            if (-not $exists) {
                $exams += [PSCustomObject]@{
                    Title = $title
                    ExamNumber = $examNumber
                    PassedDate = $passedDate
                }
            }
        }
        
        $tablePattern = "(?m)^([^\t\n\r]+?)\s{2,}([A-Z0-9-]+|[0-9]{2,3})\s{2,}(\d+\s+[A-Za-z]+\s+\d+)\s*$"
        $tableMatches = [regex]::Matches($textContent, $tablePattern)
        
        foreach ($match in $tableMatches) {
            $title = $match.Groups[1].Value.Trim()
            if (($title -notmatch "Exam title") -and 
                ($title -notmatch "^Show (more|less)$") -and
                ($title.Length -gt 3)) {
                
                $examNumber = $match.Groups[2].Value.Trim()
                $passedDate = $match.Groups[3].Value.Trim()
                
                $exists = $exams | Where-Object { $_.ExamNumber -eq $examNumber }
                
                if (-not $exists) {
                    $exams += [PSCustomObject]@{
                        Title = $title
                        ExamNumber = $examNumber
                        PassedDate = $passedDate
                    }
                }
            }
        }
    }
    
    if ($exams.Count -lt 5) {
        $multiLinePattern = "(?s)Exam title\s*\r?\n((?:.+?\r?\n)+?)Exam number\s*\r?\n([A-Z0-9-]+|[0-9]{2,3})\s*\r?\n\s*Passed date\s*\r?\n(\d+\s+[A-Za-z]+\s+\d+)"
        $multiLineMatches = [regex]::Matches($textContent, $multiLinePattern)
        
        foreach ($match in $multiLineMatches) {
            $rawTitle = $match.Groups[1].Value
            $cleanTitle = ($rawTitle -replace "\r?\n", " ").Trim()
            
            $examNumber = $match.Groups[2].Value.Trim()
            $passedDate = $match.Groups[3].Value.Trim()
            
            $exists = $exams | Where-Object { $_.ExamNumber -eq $examNumber }
            
            if (-not $exists) {
                $exams += [PSCustomObject]@{
                    Title = $cleanTitle
                    ExamNumber = $examNumber
                    PassedDate = $passedDate
                }
            }
        }
    }
    
    return $exams
}

function Extract-MctHistory {
    param (
        [string]$textContent,
        [string]$fullText
    )
    
    $mctHistory = @()
    
    if ($fullText -match "(?s)Microsoft Certified Trainer history.*?(\d+\s+[A-Za-z]+\s+\d+).*?(\d+\s+[A-Za-z]+\s+\d+)") {
        $activeFrom = $matches[1].Trim()
        $to = $matches[2].Trim()
        
        $mctHistory += [PSCustomObject]@{
            Title = "MCT History"
            ActiveFrom = $activeFrom
            To = $to
        }
        
        return $mctHistory
    }
    
    if ($textContent -match "(?s)MCT\s+title.*?Active\s+from.*?To") {
        $dateMatches = [regex]::Matches($textContent, "(\d+\s+[A-Za-z]+\s+\d+)")
        
        if ($dateMatches.Count -ge 2) {
            $activeFrom = $dateMatches[0].Value.Trim()
            $to = $dateMatches[1].Value.Trim()
            
            $mctHistory += [PSCustomObject]@{
                Title = "MCT History"
                ActiveFrom = $activeFrom
                To = $to
            }
            
            return $mctHistory
        }
    }
    
    if ($textContent -match "(?m)MCT\s+title\s+Active\s+from\s+To\s*\r?\n([^\r\n]+?)\s+(\d+\s+[A-Za-z]+\s+\d+)\s+(\d+\s+[A-Za-z]+\s+\d+)") {
        $title = $matches[1].Trim()
        $activeFrom = $matches[2].Trim()
        $to = $matches[3].Trim()
        
        $mctHistory += [PSCustomObject]@{
            Title = $title
            ActiveFrom = $activeFrom
            To = $to
        }
        
        return $mctHistory
    }
    
    if ($fullText -match "(?s)Microsoft Certified Trainer.*?Active from.*?To") {
        # Find any lines with date patterns
        $lines = $fullText -split "\r?\n"
        $dateLines = $lines | Where-Object { $_ -match "\d+\s+[A-Za-z]+\s+\d+" }
        
        if ($dateLines.Count -ge 2) {
            $activeFromMatch = [regex]::Match($dateLines[0], "\d+\s+[A-Za-z]+\s+\d+")
            $toMatch = [regex]::Match($dateLines[1], "\d+\s+[A-Za-z]+\s+\d+")
            
            if ($activeFromMatch.Success -and $toMatch.Success) {
                $activeFrom = $activeFromMatch.Value.Trim()
                $to = $toMatch.Value.Trim()
                
                $mctHistory += [PSCustomObject]@{
                    Title = "MCT History"
                    ActiveFrom = $activeFrom
                    To = $to
                }
                
                return $mctHistory
            }
        }
    }
    
    if ($fullText -like "*Active from*" -and $fullText -like "*To*") {
        $lines = $fullText -split "\r?\n"
        $activeFromLineIndex = [array]::FindIndex($lines, [Predicate[string]]{ param($line) $line -like "*Active from*" })
        $toLineIndex = [array]::FindIndex($lines, [Predicate[string]]{ param($line) $line -like "*To*" })
        
        if ($activeFromLineIndex -ne -1 -and $toLineIndex -ne -1) {
            $possibleActiveFromLine = $lines[$activeFromLineIndex + 1]
            $possibleToLine = $lines[$toLineIndex + 1]
            
            $activeFromMatch = [regex]::Match($possibleActiveFromLine, "\d+\s+[A-Za-z]+\s+\d+")
            $toMatch = [regex]::Match($possibleToLine, "\d+\s+[A-Za-z]+\s+\d+")
            
            if ($activeFromMatch.Success -and $toMatch.Success) {
                $activeFrom = $activeFromMatch.Value.Trim()
                $to = $toMatch.Value.Trim()
                
                $mctHistory += [PSCustomObject]@{
                    Title = "MCT History"
                    ActiveFrom = $activeFrom
                    To = $to
                }
                
                return $mctHistory
            }
        }
    }
    
    if ($fullText -like "*MCT*" -or $fullText -like "*Microsoft Certified Trainer*") {
        $allDates = [regex]::Matches($fullText, "(\d+\s+[A-Za-z]+\s+\d+)")
        
        if ($allDates.Count -ge 2) {
            $mctIndex = $fullText.IndexOf("Microsoft Certified Trainer")
            if ($mctIndex -eq -1) { $mctIndex = $fullText.IndexOf("MCT") }
            
            if ($mctIndex -ne -1) {
                $closestDates = $allDates | 
                    Sort-Object { [Math]::Abs($_.Index - $mctIndex) } | 
                    Select-Object -First 2
                
                if ($closestDates.Count -eq 2) {
                    $date1 = [DateTime]::ParseExact($closestDates[0].Value, "d MMM yyyy", $null)
                    $date2 = [DateTime]::ParseExact($closestDates[1].Value, "d MMM yyyy", $null)
                    
                    if ($date1 -lt $date2) {
                        $activeFrom = $closestDates[0].Value
                        $to = $closestDates[1].Value
                    } else {
                        $activeFrom = $closestDates[1].Value
                        $to = $closestDates[0].Value
                    }
                    
                    $mctHistory += [PSCustomObject]@{
                        Title = "MCT History"
                        ActiveFrom = $activeFrom
                        To = $to
                    }
                    
                    return $mctHistory
                }
            }
        }
    }
    
    return $mctHistory
}

try {
    Write-Host "Microsoft Certification Transcript Report Generator" -ForegroundColor Blue
    Write-Host "=================================================" -ForegroundColor Blue
    Write-Host ""
    Write-Host "This script will convert a saved transcript text file into a nicely formatted HTML report." -ForegroundColor Cyan
    Write-Host ""
    
    # Find the transcript file in the current directory
    $transcriptFile = Get-ChildItem -Path ".\Transcript - * _ Microsoft Learn.txt" -ErrorAction SilentlyContinue
    
    if (-not $transcriptFile) {
        Write-Host "Looking for transcript file in the script directory..." -ForegroundColor Yellow
        $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
        $transcriptFile = Get-ChildItem -Path "$scriptPath\Transcript - * _ Microsoft Learn.txt" -ErrorAction SilentlyContinue
    }
    
    if (-not $transcriptFile) {
        # Allow any text file as input
        Write-Host "No specific transcript file found. Looking for any text file..." -ForegroundColor Yellow
        $transcriptFile = Get-ChildItem -Path ".\*.txt" -ErrorAction SilentlyContinue | Out-GridView -Title "Select a transcript file" -OutputMode Single
    }
    
    if (-not $transcriptFile) {
        Write-Host "No transcript file found. Please save your Microsoft Learn transcript as a text file in the same folder as this script." -ForegroundColor Red
        exit
    }
    
    if ($transcriptFile -is [array]) {
        Write-Host "Multiple transcript files found. Using the first one: $($transcriptFile[0].Name)" -ForegroundColor Yellow
        $transcriptFile = $transcriptFile[0]
    }
    
    Write-Host "Found transcript file: $($transcriptFile.Name)" -ForegroundColor Green
    Write-Host "Reading transcript file..." -ForegroundColor Cyan
    
    $fullText = Get-Content -Path $transcriptFile.FullName -Raw
    
    # Extract only the relevant part starting from "Legal name:"
    $startIndex = $fullText.IndexOf("Legal name:")
    if ($startIndex -eq -1) {
        Write-Host "The transcript file does not contain the expected 'Legal name:' section." -ForegroundColor Red
        exit
    }
    
    $relevantText = $fullText.Substring($startIndex)
    
    Write-Host "Extracting information from transcript file..." -ForegroundColor Cyan
    
    $nameMatch = [regex]::Match($relevantText, 'Legal name:\s*(.+?)(?:\r?\n)')
    $name = if ($nameMatch.Success) { $nameMatch.Groups[1].Value.Trim() } else { "Not Found" }
    
    $emailMatch = [regex]::Match($relevantText, 'Email:\s*(.+?)(?:\r?\n)')
    $email = if ($emailMatch.Success) { $emailMatch.Groups[1].Value.Trim() } else { "Not Found" }
    
    $certIdMatch = [regex]::Match($relevantText, 'Microsoft Certification ID:\s*(.+?)(?:\r?\n)')
    $certId = if ($certIdMatch.Success) { $certIdMatch.Groups[1].Value.Trim() } else { "Not Found" }
    
    if ($certId -eq "Not Found") {
        $usernameMatch = [regex]::Match($relevantText, 'Username:\s*(.+?)(?:\r?\n|\s+Edit)')
        $certId = if ($usernameMatch.Success) { $usernameMatch.Groups[1].Value.Trim() } else { "Not Found" }
    }
    
    Write-Host "Personal Information:" -ForegroundColor Cyan
    Write-Host "  Name: $name" -ForegroundColor White
    Write-Host "  Email: $email" -ForegroundColor White
    Write-Host "  Certification ID: $certId" -ForegroundColor White
    
   
    $activeCertSection = ""
    if ($relevantText -match "(?s)Active certifications(.*?)(?:Passed exams|$)") {
        $activeCertSection = $matches[1]
    }
    
    $examSection = ""
    if ($relevantText -match "(?s)Passed exams(.*?)(?:Microsoft Certified Trainer history|$)") {
        $examSection = $matches[1]
    }
    
    if ($examSection -eq "" -or (Extract-Exams -textContent $examSection).Count -lt 20) {
        $examTablePattern = "Exam title\s+Exam number\s+Passed date\s+((?:.+\r?\n)+?)(?:Show less|\Z)"
        $examTableMatch = [regex]::Match($relevantText, $examTablePattern)
        if ($examTableMatch.Success) {
            $examSection = $examTableMatch.Value
        }
    }
    
    $mctSection = ""
    if ($relevantText -match "(?s)Microsoft Certified Trainer history(.*?)(?:Historical certifications|$)") {
        $mctSection = $matches[1]
    }
    
    if ([string]::IsNullOrWhiteSpace($mctSection) -or (Extract-MctHistory -textContent $mctSection).Count -eq 0) {
        if ($relevantText -match "(?s)MCT title\s+Active from\s+To\s+(.*?)(?:Show less|Historical certifications|$)") {
            $mctSection = $matches[1]
        }
        
         if ([string]::IsNullOrWhiteSpace($mctSection) -or (Extract-MctHistory -textContent $mctSection).Count -eq 0) {
            if ($relevantText -match "(?s)Microsoft Certified Trainer(.*?)Historical certifications") {
                $mctSection = $matches[1]
            }
        }
    }
    
    $certifications = Extract-Certifications -textContent $activeCertSection -sectionType "active"
    $passedExams = Extract-Exams -textContent $examSection
    $mctHistory = Extract-MctHistory -textContent $mctSection -fullText $relevantText
    
    $historicalCerts = Extract-HistoricalCertifications -textContent $relevantText
    
    if ($historicalCerts.Count -eq 0) {
        Write-Host "Direct historical certification extraction failed, trying fallback method..." -ForegroundColor Yellow
        
        $histSection = ""
        if ($relevantText -match "(?s)(Historical certifications.*?)Show less") {
            $histSection = $matches[1]
        } elseif ($relevantText -match "(?s)(Historical certifications.*)$") {
            $histSection = $matches[1]
        }
        
        $historicalCerts = Extract-Certifications -textContent $histSection -sectionType "historical"
    }
    
    Write-Host "Found $($certifications.Count) active certifications." -ForegroundColor $(if ($certifications.Count -gt 0) { "Green" } else { "Yellow" })
    Write-Host "Found $($passedExams.Count) passed exams." -ForegroundColor $(if ($passedExams.Count -gt 0) { "Green" } else { "Yellow" })
    
    if ($passedExams.Count -lt 20) {
        Write-Host "Attempting to extract more exams using alternative method..." -ForegroundColor Yellow
        
        if ($relevantText -match "(?s)(Passed exams.*?)Show less") {
            $examText = $matches[1]
            
            $comprehensiveExamPattern = @"
(?m)(?:^\d+\.\s*)?
Exam title\s*
([^\r\n]+?)
\s*
Exam number\s*
([A-Z0-9-]+|[0-9]{2,3})
\s*
Passed date\s*
(\d+\s+[A-Za-z]+\s+\d+)
"@
            
            $comprehensiveMatches = [regex]::Matches($examText, $comprehensiveExamPattern, [System.Text.RegularExpressions.RegexOptions]::IgnorePatternWhitespace)
            
            foreach ($match in $comprehensiveMatches) {
                $title = $match.Groups[1].Value.Trim()
                $examNumber = $match.Groups[2].Value.Trim()
                $passedDate = $match.Groups[3].Value.Trim()
                
                $exists = $passedExams | Where-Object { $_.ExamNumber -eq $examNumber }
                
                if (-not $exists) {
                    $passedExams += [PSCustomObject]@{
                        Title = $title
                        ExamNumber = $examNumber
                        PassedDate = $passedDate
                    }
                }
            }
            
            $tableFormatPattern = "(?m)^([^\r\n]+?)\s{2,}([A-Z0-9-]+|[0-9]{2,3})\s{2,}(\d+\s+[A-Za-z]+\s+\d+)\s*$"
            $tableFormatMatches = [regex]::Matches($examText, $tableFormatPattern)
            
            foreach ($match in $tableFormatMatches) {
                $title = $match.Groups[1].Value.Trim()
                if (($title -notmatch "Exam title|^Show|^$") -and 
                    ($title.Length -gt 3)) {
                    
                    $examNumber = $match.Groups[2].Value.Trim()
                    $passedDate = $match.Groups[3].Value.Trim()
                    
                    $exists = $passedExams | Where-Object { $_.ExamNumber -eq $examNumber }
                    
                    if (-not $exists) {
                        $passedExams += [PSCustomObject]@{
                            Title = $title
                            ExamNumber = $examNumber
                            PassedDate = $passedDate
                        }
                    }
                }
            }
        }
        
        Write-Host "After alternative extraction: Found $($passedExams.Count) passed exams." -ForegroundColor $(if ($passedExams.Count -gt 0) { "Green" } else { "Yellow" })
    }
    
    Write-Host "Found $($mctHistory.Count) MCT history entries." -ForegroundColor $(if ($mctHistory.Count -gt 0) { "Green" } else { "Yellow" })
    Write-Host "Found $($historicalCerts.Count) historical certifications." -ForegroundColor $(if ($historicalCerts.Count -gt 0) { "Green" } else { "Yellow" })
    
    # Generate HTML report
    Write-Host "Generating HTML report..." -ForegroundColor Cyan
    $htmlReport = Get-HtmlHeader -Title "Microsoft Certification Transcript for $name" -Name $name
    
    # Add personal information section
    $htmlReport += @"
        <div class="info-section">
            <h2>Personal Information</h2>
            <div class="info-item">
                <span class="info-label">Name:</span>
                <span>$name</span>
            </div>
            <div class="info-item">
                <span class="info-label">Email:</span>
                <span>$email</span>
            </div>
            <div class="info-item">
                <span class="info-label">Certification ID:</span>
                <span>$certId</span>
            </div>
        </div>
"@
    
    # Add stats section
    $htmlReport += @"
        <div class="info-section">
            <h2>Certification Summary</h2>
            <div class="stats-container">
                <div class="stat-box">
                    <div class="stat-number">$($certifications.Count)</div>
                    <div class="stat-label">Active Certifications</div>
                </div>
                <div class="stat-box">
                    <div class="stat-number">$($passedExams.Count)</div>
                    <div class="stat-label">Passed Exams</div>
                </div>
"@

    if ($passedExams.Count -gt 0) {
        try {
            $sortedExams = $passedExams | Sort-Object -Property { 
                try {
                    [DateTime]::ParseExact($_.PassedDate, "d MMM yyyy", $null)
                } catch {
                    Get-Date
                }
            }
            
            $earliestExamDate = $sortedExams[0].PassedDate
            
            try {
                $examDate = [DateTime]::ParseExact($earliestExamDate, "d MMM yyyy", $null)
                $yearsExperience = [Math]::Round((New-TimeSpan -Start $examDate -End (Get-Date)).TotalDays / 365, 1)
                
                $htmlReport += @"
                <div class="stat-box">
                    <div class="stat-number">$yearsExperience</div>
                    <div class="stat-label">Years Experience</div>
                </div>
"@
            } catch {
                $htmlReport += @"
                <div class="stat-box">
                    <div class="stat-number">N/A</div>
                    <div class="stat-label">Years Experience</div>
                </div>
"@
            }
        } catch {
            $htmlReport += @"
                <div class="stat-box">
                    <div class="stat-number">N/A</div>
                    <div class="stat-label">Years Experience</div>
                </div>
"@
        }
    } else {
        $htmlReport += @"
                <div class="stat-box">
                    <div class="stat-number">N/A</div>
                    <div class="stat-label">Years Experience</div>
                </div>
"@
    }
    
$mctStatusMatch = [regex]::Match($relevantText, 'MCT status\s*(\w+)')
if ($mctStatusMatch.Success) {
    $mctStatus = $mctStatusMatch.Groups[1].Value.Trim()
    # Correct spelling of "InActive" to "Inactive"
    if ($mctStatus -eq "InActive") {
        $mctStatus = "Inactive"
    }
    
    $htmlReport += @"
                <div class="stat-box">
                    <div class="stat-label"><strong>MCT Status</strong></div>
                    <div class="stat-number">$mctStatus</div>
                </div>
"@
}

$htmlReport += @"
            </div>
        </div>
"@
    
    $htmlReport += @"
        <div class="info-section">
            <h2>Active Certifications</h2>
"@
    
    if ($certifications.Count -gt 0) {
        $htmlReport += @"
            <table>
                <tr>
                    <th>Certification Title</th>
                    <th>Certification Number</th>
                    <th>Earned On</th>
                    <th>Expires On</th>
                </tr>
"@
        
        foreach ($cert in $certifications) {
            $htmlReport += @"
                <tr>
                    <td>$(Clean-CertificationTitle $cert.Title)</td>
                    <td>$($cert.CertificationNumber)</td>
                    <td>$($cert.EarnedOn)</td>
                    <td>$($cert.ExpiresOn)</td>
                </tr>
"@
        }
        
        $htmlReport += "</table>"
    } else {
        $htmlReport += "<p>No active certifications found.</p>"
    }
    
    $htmlReport += "</div>"
    
    $htmlReport += @"
    <div class="info-section">
        <h2>Passed Exams</h2>
"@

if ($passedExams.Count -gt 0) {
$htmlReport += @"
        <table>
            <tr>
                <th>Exam Title</th>
                <th>Exam Number</th>
                <th>Passed Date</th>
            </tr>
"@
    
foreach ($exam in $passedExams) {
    $htmlReport += @"
        <tr>
            <td>$(Clean-CertificationTitle $exam.Title)</td>
            <td>$($exam.ExamNumber)</td>
            <td>$($exam.PassedDate)</td>
        </tr>
"@
}
    
$htmlReport += "</table>"

$htmlReport += @"
    <p><strong>Total exams: $($passedExams.Count)</strong></p>
"@
} else {
$htmlReport += "<p>No passed exams found.</p>"
}

$htmlReport += "</div>"
    
    $htmlReport += @"
        <div class="info-section">
            <h2>Microsoft Certified Trainer (MCT) History</h2>
"@
    
    if ($mctHistory.Count -gt 0) {
        $htmlReport += @"
            <table>
                <tr>
                    <th>MCT Title</th>
                    <th>Active From</th>
                    <th>To</th>
                </tr>
"@
        
        foreach ($mct in $mctHistory) {
            # Clean up any trailing whitespace or special characters
            $title = $mct.Title.Trim()
            $activeFrom = $mct.ActiveFrom.Trim()
            $to = $mct.To.Trim()
            
            $htmlReport += @"
                <tr>
                    <td>$title</td>
                    <td>$activeFrom</td>
                    <td>$to</td>
                </tr>
"@
        }
        
        $htmlReport += "</table>"
    } else {
        Write-Host "MCT history extraction failed, trying one more comprehensive approach..." -ForegroundColor Yellow
        
        $mctSection = ""
        if ($relevantText -match "(?s)(Microsoft Certified Trainer.*?Historical certifications)") {
            $mctSection = $matches[1]
            
            $dateMatches = [regex]::Matches($mctSection, "(\d+\s+[A-Za-z]+\s+\d+)")
            
            if ($dateMatches.Count -ge 2) {
                $htmlReport += @"
                <table>
                    <tr>
                        <th>MCT Title</th>
                        <th>Active From</th>
                        <th>To</th>
                    </tr>
                    <tr>
                        <td>MCT History</td>
                        <td>$($dateMatches[0].Value.Trim())</td>
                        <td>$($dateMatches[1].Value.Trim())</td>
                    </tr>
                </table>
"@
            } else {
                $htmlReport += "<p>No MCT history found.</p>"
            }
        } else {
            $htmlReport += "<p>No MCT history found.</p>"
        }
    }
    
    $htmlReport += "</div>"
    
    $htmlReport += @"
        <div class="info-section">
            <h2>Historical Certifications</h2>
"@
    
    if ($historicalCerts.Count -gt 0) {
        $htmlReport += @"
            <table>
                <tr>
                    <th>Certification Title</th>
                    <th>Certification Number</th>
                    <th>State</th>
                    <th>Earned On</th>
                    <th>Expired On</th>
                </tr>
"@
        
        foreach ($cert in $historicalCerts) {
            $htmlReport += @"
                <tr>
                    <td>$(Clean-CertificationTitle $cert.Title)</td>
                    <td>$($cert.CertificationNumber)</td>
                    <td>$($cert.State)</td>
                    <td>$($cert.EarnedOn)</td>
                    <td>$($cert.ExpiredOn)</td>
                </tr>
"@
        }
        
        $htmlReport += "</table>"
    } else {
        $htmlReport += "<p>No historical certifications found.</p>"
    }
    
    $htmlReport += "</div>"
    
    $htmlReport += Get-HtmlFooter
    
    $outputFile = "$($transcriptFile.DirectoryName)\MicrosoftCertificationReport_$([DateTime]::Now.ToString('yyyyMMdd_HHmmss')).html"
    $htmlReport | Out-File -FilePath $outputFile -Encoding utf8
    
    Write-Host ""
    Write-Host "Report generated successfully!" -ForegroundColor Green
    Write-Host "Report saved to: $outputFile" -ForegroundColor Green
    
    $openReport = Read-Host "Would you like to open the report now? (Y/N)"
    if ($openReport -eq "Y" -or $openReport -eq "y") {
        Start-Process $outputFile
    }
    
} catch {
    Write-Error "An error occurred: $_"
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    Write-Host "Script execution failed. Please check your input and try again." -ForegroundColor Red
}