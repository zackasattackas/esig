# esig
A simple command line tool for creating email signatures in Outlook.

Using a template, esig will pull details about the current user from ADSI and insert these into the signature using find and replace. To customize these variables, the source code can be modified and recompiled to suit your needs.

Currently, the strings "{DisplayName}" and "{Title}" are replaced with the corrrepsonding attributes pulled from AD.

## Command Line Syntax

Type "esig -?" at a console to display detailed usage information.

### Example: Import a signature and apply it to new emails only.
```batch
esig.exe -Template C:\temp\MyCompanySignature.docx -NewEmails
```
### Example: Import a signature and apply it to all emails. Also configure the registry to prevent the user from changing the signature in Outlook 2016.
```batch
esig.exe -Template C:\temp\MyCompanySignature.docx -Persist -16
```
