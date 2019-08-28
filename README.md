# CreateCodeWithCode
Creates Model, View, Presenter for selected tables in an Excel file.
Excel files that have tables and have not previously been editied by the code creator
have models, a presenter, and a dialog created as well as a documentation module.

The plan is to migrate into an Add-In and make more robust.

There are two files in the zip file.
1) The code creator: CreateCodeWithCode.xlsm
2) The sample "proper" file to be processed
   A "proper" file is one that has at least one worksheet having at least one table.
The code creator has a ribbon interface with a "Go!" button to kick off the application
The user form presented will show a list of all "proper" workbooks currently open (a Browse capability is planned)
Once a workbook is selected all worksheets in that workbook that have at least one table are presented
Selecting a worksheet allows any or all tables in the workbook to be selected. Only selected tables are processed.

The result are Models, a Presenter, a Dialog (UserForm), and supporting code modules. A "documentation" module is also created that summarizes what was created. The presence of the documentation module in an otherwise proper file disqualifies that file from being processed again.

A SaveAsCopy is done using "(BKUP)" in the filename to preserve the original workbook.
