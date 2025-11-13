```markdown
# src/vba/project.md

Index of exported VBA modules and responsibilities.

Example entries:
- modInventoryPosting.bas — posting inventory transactions from sales/purchase documents; writes to tbl_InventoryTransactions.
- modPostingHelpers.bas — helper routines: AssignNextID, ColumnExistsInTable.
- modExportStructure.bas — helper to export workbook tables and named ranges to CSV/JSON.

When exporting modules:
- Use the exact VBA module name for the filename (e.g., modInventoryPosting.bas).
- Keep one module per file to enable diffs and code reviews.
```