# Submittal Register Beta

## Enhanced Document Management System

This is the **Beta version** of the Submittal Register system with advanced features including priority management and assignee filtering.

## 🚀 New Beta Features

### High Priority Management
- **Excel Column**: "High Priority(Y/N)" in Column A
- **Values**: "Y" for high priority, "N" for normal priority
- **Visual Indicators**: Red priority badges and color coding
- **Filtering**: Filter by priority level

### Assignee Tracking
- **Excel Column**: "Assignee" column for team member assignment
- **Dropdown Filter**: Filter submittals by assigned team member
- **Statistics**: Assignee workload distribution
- **Organization**: Track responsibility and workload

### Enhanced Filtering
- **Priority Filter**: High Priority / Normal Priority / All
- **Assignee Filter**: Dropdown with all unique assignees
- **Worksheet Filter**: Filter by worksheet (Scheduler, Precon, Mech, etc.)
- **Status Filter**: Filter by status (AN, NCR, Submitted, etc.)
- **Combined Filters**: Use multiple filters simultaneously

## 📋 Excel File Structure (Beta)

### Required Column Order:
1. **High Priority(Y/N)** - Enter "Y" for high priority, "N" for normal
2. **Submittal No.** - Submittal number (e.g., 01756-001.0)
3. **Spec Section Number** - Specification section number
4. **Specification Title** - Title of the specification
5. **Submittal Title** - Title of the submittal
6. **Assignee** - Team member responsible (for filtering)
7. **Parag No** - Paragraph number
8. **Description** - Submittal description
9. **Status** - Current status (AN, NCR, Submitted, etc.)
10. **Date Submitted** - Submission date
11. **Response Date** - Response date
12. **NLT Date** - No Later Than date

## 🛠️ How to Use

### 1. Update Your Excel File
- Add "High Priority(Y/N)" as Column A
- Add "Assignee" column
- Move all other columns to the right by one position
- Fill in priority values (Y/N) and assignee names

### 2. Run the Beta Script
```powershell
# Navigate to the Beta project directory
cd "Submittal Register Beta"

# Run the enhanced PowerShell script
.\Scripts\analyze_submittal_register_beta.ps1
```

### 3. View the Enhanced Report
- Open `Output/Submittal_Register_Report_Beta.html`
- Use the new filtering options
- View priority indicators and assignee statistics

## 📁 Project Structure

```
Submittal Register Beta/
├── Scripts/
│   └── analyze_submittal_register_beta.ps1
├── Templates/
├── Output/
│   └── Submittal_Register_Report_Beta.html
├── PDFs/
│   ├── AC/
│   ├── AN/
│   ├── AS/
│   ├── L/
│   ├── NCR/
│   ├── NR/
│   ├── RA/
│   ├── RC/
│   ├── REV/
│   ├── Submitted/
│   └── in progress/
├── index.html
├── upload.html
└── README.md
```

## 🎯 Key Improvements

### Visual Enhancements
- **Priority Badges**: Red "HIGH PRIORITY" badges for Y values
- **Color Coding**: High priority items have red borders and highlights
- **Enhanced Cards**: Better layout with priority indicators
- **Status Indicators**: Improved status badges and response indicators

### Analytics & Statistics
- **Priority Distribution**: Count of high vs normal priority items
- **Assignee Workload**: Distribution of submittals by assignee
- **Enhanced Dashboard**: More detailed statistics and metrics
- **Filter Results**: Real-time filtering with result counts

### User Experience
- **Advanced Filtering**: Multiple filter combinations
- **Responsive Design**: Works on all devices
- **Beta Indicators**: Clear beta version branding
- **Enhanced Navigation**: Better organization and flow

## 🔧 Technical Details

### PowerShell Script Features
- **Priority Detection**: Reads High Priority(Y/N) column
- **Assignee Extraction**: Collects unique assignees for filtering
- **Enhanced Statistics**: Priority and assignee-based analytics
- **Improved HTML Generation**: Better structure and styling

### HTML Report Features
- **Dynamic Filtering**: JavaScript-based real-time filtering
- **Priority Styling**: CSS classes for high/normal priority
- **Responsive Grid**: Adaptive layout for different screen sizes
- **Enhanced Interactivity**: Better user experience

## 🚨 Important Notes

### Excel File Requirements
- **Column A**: Must be "High Priority(Y/N)"
- **Values**: Only "Y" or "N" (case sensitive)
- **Assignee Column**: Required for filtering functionality
- **All Other Columns**: Must be shifted right by one position

### Compatibility
- **Excel Format**: .xlsx files only
- **PowerShell**: Requires PowerShell 5.1 or later
- **Browsers**: Modern browsers with JavaScript support
- **File Paths**: Update paths in script if needed

## 🔄 Migration from Original

### Steps to Migrate
1. **Backup Original**: Keep your original system as backup
2. **Update Excel**: Add the new columns as specified
3. **Test Beta**: Run the beta script and test functionality
4. **Update Paths**: Modify script paths if needed
5. **Deploy**: Use the beta system for enhanced features

### Rollback Plan
- Keep original scripts and files
- Beta system is separate and doesn't affect original
- Can switch back to original system anytime

## 📞 Support

### Troubleshooting
- **Excel Issues**: Ensure column order matches requirements
- **Script Errors**: Check file paths and Excel file format
- **Filter Issues**: Verify assignee names are consistent
- **Display Problems**: Check browser JavaScript support

### Common Issues
1. **Priority Not Showing**: Check Excel column A has "Y" or "N" values
2. **Assignee Filter Empty**: Ensure assignee column has data
3. **Script Fails**: Verify Excel file is closed and accessible
4. **HTML Not Loading**: Check file paths and permissions

## 🎉 Benefits

### For Project Managers
- **Priority Visibility**: Instantly see high priority items
- **Workload Tracking**: Monitor assignee workload distribution
- **Better Organization**: Enhanced filtering and sorting
- **Improved Efficiency**: Faster access to critical information

### For Team Members
- **Clear Assignments**: See who's responsible for what
- **Priority Focus**: Focus on high priority items first
- **Better Filtering**: Find relevant submittals quickly
- **Enhanced Interface**: More intuitive and user-friendly

---

**Submittal Register Beta** - Enhanced Document Management with Priority & Assignee Features

