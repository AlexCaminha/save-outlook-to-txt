# Save Outlook to TXT

Outlook VBA macros that export emails (and their attachments) as plain-text files to a local directory. Supports both automatic saving of incoming mail and bulk export of entire mailboxes.

I have created these macros to build NLP functionalities from emails and published as these may be useful for other objectives.

If you want to export to other formats check [MailItem.SaveAs method (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.saveas).

The output format serves me well but there may be other useful file naming conventions. Issue a PR if you want more flexibility.

## Features

- **Auto-save incoming emails** — new messages are saved to disk as they arrive via the `NewMailEx` event
- **Bulk export** — export every email across all accounts/mailboxes in a single run
- **Attachments** — saves file attachments alongside the `.txt` export
- **Safe filenames** — sanitizes subjects and folder names to remove invalid characters
- **Duplicate detection** — skips files that have already been exported
- **Path-length handling** — truncates filenames to stay within the Windows 260-character path limit
- **Multi-account** — each account is stored in its <Account Name> folder

## Project Structure

| File | Description |
|---|---|
| `SaveEmailAsTXT.bas` | Core `SaveMailAsTXT` function and `CleanFileName` helper |
| `ExportEntireMailbox.bas` | `ExportEntireMailbox` macro — recursively walks all accounts and folders |
| `NewMailEx` | `ThisOutlookSession` code — auto-exports emails on arrival |
| `TestExport.bas` | Quick diagnostic that lists all mailbox folders in the Immediate Window (**Ctrl+G**) |

## Setup

1. Open Outlook and press **Alt+F11** to open the VBA editor.
2. Import the `.bas` files:
   - **File > Import File** and select `SaveEmailAsTXT.bas` and `ExportEntireMailbox.bas`.
   - Optionally import `TestExport.bas` for diagnostics.
3. Paste the contents of the `NewMailEx` file into **ThisOutlookSession** (in the Project Explorer under *Microsoft Outlook Objects*).
4. Update the `BASE_PATH` constant in `ExportEntireMailbox.bas` and `NewMailEx` if you want to change the output directory (default: `D:\tempdata\email`).

## Usage

### Auto-save (incoming mail)

Once the `NewMailEx` code is in `ThisOutlookSession`, every new email is automatically saved to:

```
<BASE_PATH>\<Account Name>\<Folder Name>\<timestamp> - <subject>.txt
```

No action needed — it runs in the background whenever Outlook receives mail. If you want to check its functionality in real-time send a test email while you keep the VBA editor open and watch diagnostics in the Immediate Window (**Ctrl+G**). 

### Bulk export (all mailboxes)

1. Press **Alt+F11** to open the VBA editor.
2. Open the Immediate Window with **Ctrl+G**.
3. Run the `ExportEntireMailbox` macro (press **F5** or use **Run > Run Sub**).

Progress is shown in the Outlook status bar. The folder structure on disk mirrors your mailbox hierarchy.

If you run bulk export more than once the existing files will be ignored. If you need a fresh version you need to delete files and/or folders.

### Test / diagnostics

Run `TestExport` to verify that VBA can access your mailbox namespace. Output appears in the Immediate Window (**Ctrl+G**).

## Output Format

Emails are saved as plain-text `.txt` files. Filenames follow the pattern:

```
yyyy-MM-dd HHmmss - Subject.txt
```

Attachments are saved alongside the email:

```
yyyy-MM-dd HHmmss - Subject - attachment.ext
```

## Requirements

- Microsoft Outlook Classic (desktop, Windows)
- Macros must be enabled (File > Options > Trust Center > Macro Settings)
