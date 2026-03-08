# Quick Start

## Files you need

- `template/mail_template.docx`
- `example/sample_data.xlsx`
- `src/WordBatchMailSender.bas`

## Basic workflow

1. Open the Word template.
2. Connect it to the Excel data source.
3. Replace placeholder text with mail merge fields.
4. Import `src/WordBatchMailSender.bas` into the VBA editor.
5. Run `BatchMailMergeSend` in Word.
6. Choose the Outlook account, subject, and send mode.

## Send modes

- `1` = direct send
- `2` = save as draft

## Recommended columns

- `Email_Addresses` or similar recipient field
- `Attachment` or similar file path field
- Additional merge fields such as name, deadline, notice content, etc.
