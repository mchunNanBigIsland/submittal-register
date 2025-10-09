# Excel File Structure Guide for Submittal Register Beta

## Required Column Structure

Your Excel file must have the following columns in this exact order:

| Column | Header | Description | Example Values |
|--------|--------|-------------|----------------|
| A | High Priority(Y/N) | Priority level | Y, N |
| B | Submittal No. | Submittal number | 01756-001.0 |
| C | Spec Section Number | Specification section | 01756 |
| D | Specification Title | Title of specification | COMMISSIONING |
| E | Submittal Title | Title of submittal | Contractor Designated Training Coordinator |
| F | Assignee | Team member responsible | John Smith |
| G | Parag No | Paragraph number | 3.13-B |
| H | Description | Submittal description | Submit qualifications |
| I | Status | Current status | AN, NCR, Submitted |
| J | Date Submitted | Submission date | 05/28/2025 |
| K | Response Date | Response date | 08/14/2025 |
| L | NLT Date | No Later Than date | 10/15/2025 |

## Priority Column (Column A) - High Priority(Y/N)

### Values:
- **Y** = High Priority (will show red "HIGH PRIORITY" badge)
- **N** = Normal Priority (no special badge)
- **Empty** = Treated as Normal Priority

### When to Use High Priority (Y):
- Critical path items
- Items blocking other work
- Items with tight deadlines
- Items requiring immediate attention
- Items that are overdue or approaching deadline

### When to Use Normal Priority (N):
- Standard processing items
- Items with flexible timelines
- Routine submittals
- Items not on critical path

## Assignee Column (Column F)

### Purpose:
- Track responsibility for each submittal
- Enable filtering by team member
- Monitor workload distribution
- Improve accountability

### Best Practices:
- Use consistent names (e.g., "John Smith" not "J. Smith")
- Include full names for clarity
- Use same format throughout file
- Consider using initials if space is limited

### Examples:
- John Smith
- Sarah Johnson
- Mike Chen
- Lisa Rodriguez

## Status Column (Column I)

### Valid Status Values:
- **AN** - Approved No Exceptions
- **AC** - Approved with Comments
- **AS** - Approved Correct and Resubmit
- **NCR** - Not Approved Correct Resubmit
- **NR** - Not Approved Rejected
- **RA** - Receipt Acknowledged
- **RC** - Receipt with Comments
- **L** - Living Document
- **REV** - Revision
- **Submitted** - Submitted
- **in progress** - In Progress

## Date Columns

### Date Submitted (Column J):
- Format: MM/dd/yyyy (e.g., 05/28/2025)
- Use for calculating response due dates
- Leave empty if not submitted yet

### Response Date (Column K):
- Format: MM/dd/yyyy (e.g., 08/14/2025)
- Date when response was received
- Leave empty if no response yet

### NLT Date (Column L):
- Format: MM/dd/yyyy (e.g., 10/15/2025)
- "No Later Than" date
- Critical deadline for the submittal

## Sample Data Row

| High Priority(Y/N) | Submittal No. | Spec Section Number | Specification Title | Submittal Title | Assignee | Parag No | Description | Status | Date Submitted | Response Date | NLT Date |
|-------------------|---------------|-------------------|-------------------|----------------|----------|----------|-------------|--------|----------------|---------------|----------|
| Y | 01756-001.0 | 01756 | COMMISSIONING | Contractor Designated Training Coordinator | John Smith | 3.13-B | Submit qualifications | NCR | 05/28/2025 | 08/14/2025 | 10/15/2025 |
| N | 01756-002.0 | 01756 | COMMISSIONING | Training Program | Sarah Johnson | 1.05-A | Submit training program | AN | 06/01/2025 | 06/15/2025 | 07/01/2025 |

## Migration from Original Excel

### Step 1: Add New Columns
1. Insert new column A for "High Priority(Y/N)"
2. Insert new column F for "Assignee"
3. All existing columns shift right by one position

### Step 2: Fill Priority Values
1. Review each submittal
2. Mark critical items as "Y"
3. Mark normal items as "N"
4. Leave empty for normal priority

### Step 3: Add Assignee Information
1. Assign team members to each submittal
2. Use consistent naming format
3. Consider workload distribution
4. Update as assignments change

### Step 4: Test the Beta System
1. Run the beta PowerShell script
2. Check that priorities show correctly
3. Verify assignee filtering works
4. Test all filter combinations

## Common Issues and Solutions

### Issue: Priority not showing
**Solution**: Check that Column A has "Y" or "N" values (case sensitive)

### Issue: Assignee filter empty
**Solution**: Ensure Column F has assignee names, not empty cells

### Issue: Script fails to read Excel
**Solution**: 
- Close Excel file before running script
- Check file path in script
- Ensure file is .xlsx format

### Issue: Dates not formatting correctly
**Solution**: Use MM/dd/yyyy format in Excel cells

### Issue: Status not recognized
**Solution**: Use exact status values from the list above

## Best Practices

### Priority Management
- Review priorities regularly
- Update as project progresses
- Use Y sparingly for true high priority items
- Consider project timeline and dependencies

### Assignee Management
- Assign based on expertise and availability
- Update assignments as needed
- Consider workload balance
- Use consistent naming

### Data Quality
- Keep all columns filled where possible
- Use consistent formatting
- Regular data validation
- Backup your Excel file

## Excel Template

You can create a template with the proper column headers:

1. **Row 1**: Leave empty or add project information
2. **Row 2**: Column headers (as listed above)
3. **Row 3+**: Data rows

### Header Row Example:
```
High Priority(Y/N) | Submittal No. | Spec Section Number | Specification Title | Submittal Title | Assignee | Parag No | Description | Status | Date Submitted | Response Date | NLT Date
```

This structure ensures compatibility with the Submittal Register Beta system and enables all the enhanced features.

