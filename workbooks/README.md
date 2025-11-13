```markdown
# workbooks/

Purpose:
- Store sample and original Excel workbooks used for development, testing and debugging.

Structure:
- originals/: Raw .xlsm/.xlsx files (use Git LFS for large or sensitive files).
- samples/: Redacted, small example workbooks safe for public sharing.

Guidelines:
- Do not commit production PII or sensitive data. Create a redacted copy before committing to samples/.
- Export VBA modules to src/vba/modules/ to keep code reviewable and diffable.
- Run the export_structure macro (if available in tools/) to generate docs/reports/{workbook}_structure.csv and commit those reports.
```