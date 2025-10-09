# Submittal Register Beta Analysis Script
# Enhanced version with High Priority (Y/N) column and Assignee filtering
# This script reads the Excel submittal register and generates a comprehensive HTML report

# Function to convert Excel date to proper format
function Convert-ExcelDate {
    param($excelDate)
    if ($excelDate -eq $null -or $excelDate -eq "") { return "" }
    
    # Handle "N/a" or "n/a" values
    if ($excelDate.ToString().ToLower() -eq "n/a") { return "N/A" }
    
    # Check if it's a number (Excel serial date)
    if ($excelDate -match '^\d+$' -or ($excelDate -is [double] -or $excelDate -is [int])) {
        try {
            $serialDate = [double]$excelDate
            if ($serialDate -gt 0) {
                $date = [DateTime]::FromOADate($serialDate)
                return $date.ToString("MM/dd/yyyy")
            }
        } catch {
            return $excelDate.ToString()
        }
    }
    
    # If it's already a date string, try to parse and format it
    try {
        $parsedDate = [DateTime]::Parse($excelDate.ToString())
        return $parsedDate.ToString("MM/dd/yyyy")
    } catch {
        return $excelDate.ToString()
    }
}

# Function to get priority class and styling
function Get-PriorityInfo {
    param($priorityValue)
    
    if ($priorityValue -eq "Y" -or $priorityValue -eq "Yes" -or $priorityValue -eq "YES") {
        return @{
            Class = "high-priority"
            Badge = "HIGH PRIORITY"
            Color = "#ff4444"
            Icon = "🔴"
        }
    } elseif ($priorityValue -eq "N" -or $priorityValue -eq "No" -or $priorityValue -eq "NO") {
        return @{
            Class = "normal-priority"
            Badge = ""
            Color = "#666666"
            Icon = ""
        }
    } else {
        return @{
            Class = "normal-priority"
            Badge = ""
            Color = "#666666"
            Icon = ""
        }
    }
}

# Function to calculate response date and urgency for submitted submittals
function Get-ResponseInfo {
    param($submissionDate, $currentStatus)
    
    # Only calculate for submitted status
    if ($currentStatus -ne "Submitted") {
        return @{
            ResponseDueDate = ""
            DaysOverdue = 0
            UrgencyLevel = "none"
            UrgencyClass = ""
            ResponseStatus = ""
        }
    }
    
    # Parse submission date
    $submissionDateTime = $null
    if ($submissionDate -and $submissionDate -ne "" -and $submissionDate -ne "N/A") {
        try {
            $submissionDateTime = [DateTime]::Parse($submissionDate)
        } catch {
            return @{
                ResponseDueDate = ""
                DaysOverdue = 0
                UrgencyLevel = "none"
                UrgencyClass = ""
                ResponseStatus = ""
            }
        }
    } else {
        return @{
            ResponseDueDate = ""
            DaysOverdue = 0
            UrgencyLevel = "none"
            UrgencyClass = ""
            ResponseStatus = ""
        }
    }
    
    # Calculate response due date (30 days from submission)
    $responseDueDate = $submissionDateTime.AddDays(30)
    $currentDate = Get-Date
    $daysOverdue = ($currentDate - $responseDueDate).Days
    
    # Determine urgency level
    $urgencyLevel = "none"
    $urgencyClass = ""
    $responseStatus = ""
    
    if ($daysOverdue > 0) {
        $urgencyLevel = "overdue"
        $urgencyClass = "overdue"
        $responseStatus = "LATE"
    } elseif ($daysOverdue >= -3) {
        $urgencyLevel = "approaching"
        $urgencyClass = "approaching"
        $responseStatus = "Due Soon"
    } else {
        $urgencyLevel = "on-time"
        $urgencyClass = "on-time"
        $responseStatus = "On Time"
    }
    
    return @{
        ResponseDueDate = $responseDueDate.ToString("MM/dd/yyyy")
        DaysOverdue = $daysOverdue
        UrgencyLevel = $urgencyLevel
        UrgencyClass = $urgencyClass
        ResponseStatus = $responseStatus
    }
}

# Function to get unique assignees for dropdown
function Get-UniqueAssignees {
    param($allSubmittals)
    
    $assignees = @()
    foreach ($submittal in $allSubmittals) {
        if ($submittal.Assignee -and $submittal.Assignee.Trim() -ne "" -and $submittal.Assignee -ne "N/A") {
            if ($assignees -notcontains $submittal.Assignee.Trim()) {
                $assignees += $submittal.Assignee.Trim()
            }
        }
    }
    return $assignees | Sort-Object
}

try {
    # Load Excel file
    $excelFile = "Official_Submittal_Register_Updated.xlsx"
    Write-Host "Processing Excel file: $excelFile" -ForegroundColor Yellow
    
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Open workbook
    $workbook = $excel.Workbooks.Open((Resolve-Path $excelFile).Path)
    
    # Get all worksheets
    $worksheets = $workbook.Worksheets
    $allSubmittals = @()
    $worksheetNames = @()
    
    # Process each worksheet
    foreach ($worksheet in $worksheets) {
        $worksheetName = $worksheet.Name
        Write-Host "Processing worksheet: $worksheetName" -ForegroundColor Yellow
        
        if ($worksheetName -notin @("Summary", "Dashboard", "Charts")) {
            $worksheetNames += $worksheetName
            
            # Get used range
            $usedRange = $worksheet.UsedRange
            if ($usedRange -and $usedRange.Rows.Count -gt 1) {
                # Get headers (assuming row 2 has headers based on your Excel structure)
                $headers = @()
                for ($col = 1; $col -le $usedRange.Columns.Count; $col++) {
                    $headerValue = $usedRange.Cells.Item(2, $col).Value2
                    if ($headerValue) {
                        $headers += $headerValue.ToString().Trim()
                    } else {
                        $headers += ""
                    }
                }
                
                # Process data rows (starting from row 3)
                for ($row = 3; $row -le $usedRange.Rows.Count; $row++) {
                    $submittalData = @{}
                    
                    # Map data to headers
                    for ($col = 1; $col -le $usedRange.Columns.Count; $col++) {
                        $cellValue = $usedRange.Cells.Item($row, $col).Value2
                        if ($cellValue) {
                            $submittalData[$headers[$col-1]] = $cellValue.ToString().Trim()
                        } else {
                            $submittalData[$headers[$col-1]] = ""
                        }
                    }
                    
                    # Only process rows with submittal numbers
                    if ($submittalData["Submittal No."] -and $submittalData["Submittal No."] -ne "") {
                        $submittal = @{
                            SubmittalNumber = $submittalData["Submittal No."]
                            SpecSectionNumber = $submittalData["Spec Section Number"]
                            SpecificationTitle = $submittalData["Specification Title"]
                            SubmittalTitle = $submittalData["Submittal Title"]
                            ParagraphNumber = $submittalData["Parag No"]
                            Description = $submittalData["Description"]
                            Assignee = $submittalData["Assignee"]
                            Worksheet = $worksheetName
                            HighPriority = $submittalData["High Priority(Y/N)"]
                            Status = $submittalData["Status"]
                            DateSubmitted = Convert-ExcelDate $submittalData["Date Submitted"]
                            ResponseDate = Convert-ExcelDate $submittalData["Response Date"]
                            NLTDate = Convert-ExcelDate $submittalData["NLT Date"]
                        }
                        
                        # Get priority information
                        $priorityInfo = Get-PriorityInfo $submittal.HighPriority
                        $submittal.PriorityClass = $priorityInfo.Class
                        $submittal.PriorityBadge = $priorityInfo.Badge
                        $submittal.PriorityColor = $priorityInfo.Color
                        $submittal.PriorityIcon = $priorityInfo.Icon
                        
                        # Get response information for submitted items
                        $responseInfo = Get-ResponseInfo $submittal.DateSubmitted $submittal.Status
                        $submittal.ResponseDueDate = $responseInfo.ResponseDueDate
                        $submittal.DaysOverdue = $responseInfo.DaysOverdue
                        $submittal.UrgencyLevel = $responseInfo.UrgencyLevel
                        $submittal.UrgencyClass = $responseInfo.UrgencyClass
                        $submittal.ResponseStatus = $responseInfo.ResponseStatus
                        
                        $allSubmittals += $submittal
                    }
                }
            }
        }
    }
    
    # Get unique assignees
    $uniqueAssignees = Get-UniqueAssignees $allSubmittals
    
    # Calculate statistics
    $totalSubmittals = $allSubmittals.Count
    $highPriorityCount = ($allSubmittals | Where-Object { $_.HighPriority -eq "Y" }).Count
    $normalPriorityCount = $totalSubmittals - $highPriorityCount
    
    # Group by status
    $statusGroups = $allSubmittals | Group-Object Status
    $statusStats = @{}
    foreach ($group in $statusGroups) {
        $statusStats[$group.Name] = $group.Count
    }
    
    # Group by assignee
    $assigneeGroups = $allSubmittals | Group-Object Assignee
    $assigneeStats = @{}
    foreach ($group in $assigneeGroups) {
        if ($group.Name -and $group.Name.Trim() -ne "") {
            $assigneeStats[$group.Name] = $group.Count
        }
    }
    
    Write-Host "Total unique submittals found: $totalSubmittals" -ForegroundColor Green
    Write-Host "High priority submittals: $highPriorityCount" -ForegroundColor Red
    Write-Host "Normal priority submittals: $normalPriorityCount" -ForegroundColor Green
    Write-Host "Unique assignees: $($uniqueAssignees.Count)" -ForegroundColor Cyan
    
    # Generate HTML report
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Submittal Register Beta - Enhanced Analysis Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            min-height: 100vh;
            color: #333;
        }
        
        .header {
            background: linear-gradient(135deg, #0f4c75 0%, #1e3c72 100%);
            color: white;
            padding: 2rem 0;
            text-align: center;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        }
        
        .header h1 {
            font-size: 2.5rem;
            margin-bottom: 0.5rem;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        
        .header p {
            font-size: 1.2rem;
            opacity: 0.9;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 2rem;
        }
        
        .dashboard {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 2rem;
            margin-bottom: 3rem;
        }
        
        .card {
            background: white;
            border-radius: 15px;
            padding: 2rem;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            transition: transform 0.3s ease;
        }
        
        .card:hover {
            transform: translateY(-5px);
        }
        
        .card h3 {
            color: #1e3c72;
            margin-bottom: 1rem;
            font-size: 1.5rem;
        }
        
        .stat-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 1rem 0;
            border-bottom: 1px solid #eee;
        }
        
        .stat-item:last-child {
            border-bottom: none;
        }
        
        .stat-label {
            font-weight: 600;
            color: #555;
        }
        
        .stat-value {
            font-size: 1.2rem;
            font-weight: bold;
        }
        
        .high-priority {
            color: #ff4444;
        }
        
        .normal-priority {
            color: #666;
        }
        
        .priority-badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: bold;
            text-transform: uppercase;
            margin-left: 8px;
        }
        
        .priority-badge.high {
            background: linear-gradient(135deg, #ff4444, #cc0000);
            color: white;
            box-shadow: 0 2px 4px rgba(255, 68, 68, 0.3);
        }
        
        .priority-badge.normal {
            background: linear-gradient(135deg, #666666, #444444);
            color: white;
            box-shadow: 0 2px 4px rgba(102, 102, 102, 0.3);
        }
        
        .filters {
            background: white;
            border-radius: 15px;
            padding: 2rem;
            margin-bottom: 2rem;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        }
        
        .filter-group {
            display: flex;
            gap: 1rem;
            align-items: center;
            flex-wrap: wrap;
        }
        
        .filter-group label {
            font-weight: 600;
            color: #555;
        }
        
        .filter-group select {
            padding: 8px 12px;
            border: 2px solid #ddd;
            border-radius: 8px;
            font-size: 14px;
            min-width: 150px;
        }
        
        .filter-group select:focus {
            outline: none;
            border-color: #1e3c72;
        }
        
        .submittal-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(400px, 1fr));
            gap: 1.5rem;
        }
        
        .submittal-card {
            background: white;
            border-radius: 12px;
            padding: 1.5rem;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            transition: all 0.3s ease;
            border-left: 4px solid #ddd;
        }
        
        .submittal-card.high-priority {
            border-left-color: #ff4444;
            box-shadow: 0 5px 15px rgba(255, 68, 68, 0.2);
        }
        
        .submittal-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        }
        
        .submittal-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1rem;
        }
        
        .submittal-number {
            font-size: 1.2rem;
            font-weight: bold;
            color: #1e3c72;
        }
        
        .submittal-title {
            font-size: 1.1rem;
            font-weight: 600;
            color: #333;
            margin-bottom: 0.5rem;
        }
        
        .submittal-details {
            font-size: 0.9rem;
            color: #666;
            line-height: 1.4;
        }
        
        .assignee {
            font-weight: 600;
            color: #1e3c72;
        }
        
        .worksheet {
            background: #f0f8ff;
            color: #1e3c72;
            padding: 4px 8px;
            border-radius: 6px;
            font-size: 0.8rem;
            font-weight: 600;
        }
        
        .status-badge {
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: bold;
            text-transform: uppercase;
        }
        
        .status-badge.ncr { background: #ffebee; color: #c62828; }
        .status-badge.an { background: #e8f5e8; color: #2e7d32; }
        .status-badge.submitted { background: #e0f2f1; color: #00695c; }
        .status-badge.ra { background: #e3f2fd; color: #1565c0; }
        .status-badge.rc { background: #f3e5f5; color: #7b1fa2; }
        
        .response-status {
            font-size: 0.8rem;
            padding: 2px 8px;
            border-radius: 12px;
            font-weight: 600;
        }
        
        .response-status.overdue {
            background: #ffebee;
            color: #c62828;
        }
        
        .response-status.approaching {
            background: #fff3e0;
            color: #e65100;
        }
        
        .response-status.on-time {
            background: #e8f5e8;
            color: #2e7d32;
        }
        
        .no-results {
            text-align: center;
            padding: 3rem;
            color: #666;
            font-size: 1.1rem;
        }
        
        .beta-badge {
            background: linear-gradient(135deg, #ff6b6b, #ee5a24);
            color: white;
            padding: 8px 16px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: bold;
            text-transform: uppercase;
            margin-left: 10px;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>📋 Submittal Register Beta <span class="beta-badge">BETA</span></h1>
        <p>Enhanced Analysis Report with Priority Management & Assignee Filtering</p>
    </div>
    
    <div class="container">
        <!-- Dashboard -->
        <div class="dashboard">
            <div class="card">
                <h3>📊 Overview Statistics</h3>
                <div class="stat-item">
                    <span class="stat-label">Total Submittals</span>
                    <span class="stat-value">$totalSubmittals</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">High Priority</span>
                    <span class="stat-value high-priority">$highPriorityCount</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">Normal Priority</span>
                    <span class="stat-value normal-priority">$normalPriorityCount</span>
                </div>
            </div>
            
            <div class="card">
                <h3>👥 Assignee Distribution</h3>
"@

    # Add assignee statistics
    foreach ($assignee in $uniqueAssignees | Sort-Object) {
        $count = $assigneeStats[$assignee]
        $htmlContent += @"
                <div class="stat-item">
                    <span class="stat-label">$assignee</span>
                    <span class="stat-value">$count</span>
                </div>
"@
    }

    $htmlContent += @"
            </div>
            
            <div class="card">
                <h3>📈 Status Distribution</h3>
"@

    # Add status statistics
    foreach ($status in $statusStats.Keys | Sort-Object) {
        $count = $statusStats[$status]
        $htmlContent += @"
                <div class="stat-item">
                    <span class="stat-label">$status</span>
                    <span class="stat-value">$count</span>
                </div>
"@
    }

    $htmlContent += @"
            </div>
        </div>
        
        <!-- Filters -->
        <div class="filters">
            <h3>🔍 Filter Options</h3>
            <div class="filter-group">
                <label for="priorityFilter">Priority:</label>
                <select id="priorityFilter">
                    <option value="">All Priorities</option>
                    <option value="high">High Priority Only</option>
                    <option value="normal">Normal Priority Only</option>
                </select>
                
                <label for="assigneeFilter">Assignee:</label>
                <select id="assigneeFilter">
                    <option value="">All Assignees</option>
"@

    # Add assignee options
    foreach ($assignee in $uniqueAssignees) {
        $htmlContent += @"
                    <option value="$assignee">$assignee</option>
"@
    }

    $htmlContent += @"
                </select>
                
                <label for="worksheetFilter">Worksheet:</label>
                <select id="worksheetFilter">
                    <option value="">All Worksheets</option>
"@

    # Add worksheet options
    foreach ($worksheet in $worksheetNames) {
        $htmlContent += @"
                    <option value="$worksheet">$worksheet</option>
"@
    }

    $htmlContent += @"
                </select>
                
                <label for="statusFilter">Status:</label>
                <select id="statusFilter">
                    <option value="">All Statuses</option>
"@

    # Add status options
    foreach ($status in $statusStats.Keys | Sort-Object) {
        $htmlContent += @"
                    <option value="$status">$status</option>
"@
    }

    $htmlContent += @"
                </select>
            </div>
        </div>
        
        <!-- Submittals Grid -->
        <div class="submittal-grid" id="submittalsGrid">
"@

    # Add submittal cards
    foreach ($submittal in $allSubmittals) {
        $priorityClass = if ($submittal.HighPriority -eq "Y") { "high-priority" } else { "normal-priority" }
        $priorityBadge = if ($submittal.HighPriority -eq "Y") { '<span class="priority-badge high">HIGH PRIORITY</span>' } else { "" }
        $responseStatusHtml = if ($submittal.ResponseStatus -ne "") { "<span class='response-status $($submittal.UrgencyClass)'>$($submittal.ResponseStatus)</span>" } else { "" }
        
        $htmlContent += @"
            <div class="submittal-card $priorityClass" 
                 data-priority="$($submittal.HighPriority)" 
                 data-assignee="$($submittal.Assignee)" 
                 data-worksheet="$($submittal.Worksheet)" 
                 data-status="$($submittal.Status)">
                <div class="submittal-header">
                    <div class="submittal-number">$($submittal.SubmittalNumber)</div>
                    <div>
                        <span class="status-badge $($submittal.Status.ToLower())">$($submittal.Status)</span>
                        $priorityBadge
                    </div>
                </div>
                <div class="submittal-title">$($submittal.SubmittalTitle)</div>
                <div class="submittal-details">
                    <div><strong>Spec:</strong> $($submittal.SpecSectionNumber) - $($submittal.SpecificationTitle)</div>
                    <div><strong>Paragraph:</strong> $($submittal.ParagraphNumber)</div>
                    <div><strong>Assignee:</strong> <span class="assignee">$($submittal.Assignee)</span></div>
                    <div><strong>Worksheet:</strong> <span class="worksheet">$($submittal.Worksheet)</span></div>
                    <div><strong>Description:</strong> $($submittal.Description)</div>
                    <div><strong>Submitted:</strong> $($submittal.DateSubmitted)</div>
                    <div><strong>Response Due:</strong> $($submittal.ResponseDueDate) $responseStatusHtml</div>
                </div>
            </div>
"@
    }

    $htmlContent += @"
        </div>
        
        <div class="no-results" id="noResults" style="display: none;">
            <h3>No submittals match your current filters</h3>
            <p>Try adjusting your filter criteria to see more results.</p>
        </div>
    </div>
    
    <script>
        // Filter functionality
        function applyFilters() {
            const priorityFilter = document.getElementById('priorityFilter').value;
            const assigneeFilter = document.getElementById('assigneeFilter').value;
            const worksheetFilter = document.getElementById('worksheetFilter').value;
            const statusFilter = document.getElementById('statusFilter').value;
            
            const cards = document.querySelectorAll('.submittal-card');
            let visibleCount = 0;
            
            cards.forEach(card => {
                let show = true;
                
                // Priority filter
                if (priorityFilter === 'high' && card.dataset.priority !== 'Y') {
                    show = false;
                } else if (priorityFilter === 'normal' && card.dataset.priority !== 'N') {
                    show = false;
                }
                
                // Assignee filter
                if (assigneeFilter && card.dataset.assignee !== assigneeFilter) {
                    show = false;
                }
                
                // Worksheet filter
                if (worksheetFilter && card.dataset.worksheet !== worksheetFilter) {
                    show = false;
                }
                
                // Status filter
                if (statusFilter && card.dataset.status !== statusFilter) {
                    show = false;
                }
                
                if (show) {
                    card.style.display = 'block';
                    visibleCount++;
                } else {
                    card.style.display = 'none';
                }
            });
            
            // Show/hide no results message
            const noResults = document.getElementById('noResults');
            if (visibleCount === 0) {
                noResults.style.display = 'block';
            } else {
                noResults.style.display = 'none';
            }
        }
        
        // Add event listeners
        document.getElementById('priorityFilter').addEventListener('change', applyFilters);
        document.getElementById('assigneeFilter').addEventListener('change', applyFilters);
        document.getElementById('worksheetFilter').addEventListener('change', applyFilters);
        document.getElementById('statusFilter').addEventListener('change', applyFilters);
        
        // Initialize filters
        applyFilters();
    </script>
</body>
</html>
"@

    # Save HTML report with proper encoding (UTF-8 with BOM for better compatibility)
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText("Submittal_Register_Report_Beta.html", $htmlContent, $utf8NoBom)
    
    Write-Host "Enhanced HTML report generated successfully: Submittal_Register_Report_Beta.html" -ForegroundColor Green
    
} catch {
    Write-Error "Error reading Excel file: $($_.Exception.Message)"
} finally {
    # Clean up Excel objects
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host ""
Write-Host "Submittal Register Beta analysis completed!" -ForegroundColor Green
Write-Host "Enhanced features:" -ForegroundColor Cyan
Write-Host "  ✓ High Priority (Y/N) column support" -ForegroundColor White
Write-Host "  ✓ Assignee filtering with dropdown" -ForegroundColor White
Write-Host "  ✓ Worksheet filtering" -ForegroundColor White
Write-Host "  ✓ Priority-based visual indicators" -ForegroundColor White
Write-Host "  ✓ Enhanced statistics and dashboard" -ForegroundColor White

