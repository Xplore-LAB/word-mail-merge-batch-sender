# Architecture

```text
Excel recipient list
        |
        v
Word Mail Merge template
        |
        v
VBA module (BatchMailMergeSend)
        |
        v
Outlook account / draft / send
```

## Flow

1. Excel stores recipient addresses, attachment paths, and merge fields.
2. Word keeps the formatted mail template and merge fields.
3. VBA reads the active mail merge record row by row.
4. Outlook sends the personalized email or saves it as a draft.
5. Each row can point to a different attachment file.

## Core Value

Native Word Mail Merge is good at generating personalized bodies, but it does not naturally attach a different file for each row. This project fills that gap without replacing the existing Office workflow.
