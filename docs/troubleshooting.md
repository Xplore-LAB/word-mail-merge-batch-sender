# Troubleshooting

## Outlook account not detected

- Open Outlook first
- Make sure at least one account is signed in
- Retry the macro after Outlook finishes loading

## Attachment not added

- Check that the file path exists
- Prefer absolute Windows paths
- Confirm the attachment field name matches your Excel header

## Merge fields not replaced

- Reinsert the Word merge fields
- Confirm the Excel sheet headers match the intended names
- Avoid extra spaces in header cells

## Macro cannot run

- Save the document as a macro-enabled file when needed
- Enable Office macro permissions
- Re-import the `.bas` module if necessary

## Messages are created but not sent

- Verify which mode you selected: draft or direct send
- Check Outlook security prompts or policies
