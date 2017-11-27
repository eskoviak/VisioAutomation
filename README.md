# Welcome to the MS Office Automation Git Repository
## This repository contains PowerShell Cmdlets and Modules for use with MS Office Applications.  The purpose of these automation scripts is to externalize the automation of certain business functions.
### By externalizing these functions, the overhead of adding VBA or Macro code to the file itself is avoided--while this imbeddded approach is convenient, it has several drawbacks:
1. The type of file saved is modidied, _i.e._ MyDocument.docm vs. MyDocument.docx.  The file type suffix ending in '.xxxm' indicates that the file contains macro code.  The files are problematic to share, as security settings require the files to be digitally signed or the user to open up security on the device..
2. The automations are (mostly) used to aid the author.  Hence, they are not needed in "published" versions.
3. The above strategy may be implemented in combination with a SharePoint Major/Minor strategy.  This means that Published (Major) versions should *NOT* have embedded macros.  Minor documents, which are actively being edited, may want to use the automation.
4. Externalizing the automation routines avoids having the manage this.
5. The Published version may be readily converted to PDF format.
### Additional benefits:
1. Common routines are easier to share.
2. Automations may be easily tracked and versioned using Git.
3. No separate IDE or Compiler is necessary (as would be the case with C# or VB/VBA).