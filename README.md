# SAS_Excel_Macros
A set of SAS macros to communicate between SAS and Excel via DDE.... Or how to manipulate Excel from the comfort of your SAS code.

## Contents of this repo
* **excel_macros.sas**:  This is the SAS program that contains all of the SAS macros described in this document.
* **test_excel_macros.sas**: This is simple SAS code to demonstrate how the macros can be used. It is provided for demonstration purposes.
* **test_excel_macros.xlsx**: This is a spreadhseet that the program test_excel_macros.sas will open and manipulate. It is provided for demonstration purposes.
* **testdir\test_excel_macros_newname.xlsx**: This is the spreadsheet that test_excel_macros.sas will create.  It is provided for demonstration purposes.


## Disclaimer
I know there are better ways to do this!  DDE is old technology and is generally not supported anymore.  I highly recommend you look into using Python as a more robust and better-supported method of manipulating and filling spreadsheets.  I wrote these macros years ago, before Python was really a thing and when my only option to fill tailored spreadsheets in an automated fashion was to make SAS do it for me.  These are still useful, but updates in both SAS and Excel have broken the macros over the years.  Sometimes I can fix it.  Sometimes I abandon ship and use Python. 

## Motivation
DDE is very useful when you have formatting requirements for your spreadsheet that ODS cannot accomodate.  For example, DDE allows you to open, manipulate and populate an existing spreadsheet.  When reading information from a spreadsheet into SAS, proc import has limitations and often automatically formats data in useless or even harmful ways.  I wrote these macros after having to perform the same series of tasks over and over for various projects.  I took the most common tasks and put them into a series of macros that I can include and call from any of my programs.

## Types of macros
* Basic Worksheet Manipulations– copying, renaming, moving, and deleting worksheets.
* Obtaining workbook information– number and names of all worksheets, number of columns and rows in each worksheet.
* Reading all worksheets into datasets that can later be formatted and manipulated as needed.
* All of these are in the program excel_macros.sas which can be included in any other program.

## How it works
* What this program really does is create VBA commands and sends them to Excel
* For some of the macros, we actually create a new sheet, execute VBA in that sheet and then delete the sheet.


## Note about DDE limitations
* DDE will not tell you when it fails or runs incorrectly. This is a serious flaw in DDE.  
* Where possible I tried to add warnings and run information to the SAS lst file.  
* The warning prints try to get around this, but they do not solve the entire problem.  
* Whenever you use DDE you need to be extremely careful and double-check your results.  
* I highly recommend that after you output your data to the spreadsheet, that you read it back into SAS and run a proc compare to make sure your output matches the data you intended to output.

## List of macros and how to use them

### Macro STARTXL
* Purpose: Opens Excel system
* This macro will open Excel in a Windows platform, check that it opened correctly and will put a warning in the SAS lst file if something went wrong
* Example Call: %startxls;

### Macro NEW_XLS
* Purpose: Creates a new Excel workbook with one blank worksheet (Sheet1)
* If Excel is not open, it will open Excel before creating the blank spreadsheet.
* Example Call: %new_xls;

### Macro OPEN_XLS_FILE
* Purpose: Open a specific, existing Excel file
* If Excel is not open, it will open Excel before opening the specified workbook.
* Example Call: %open_xls_file(path=&top_folder,workbook=test_excel_macros.xlsx);
    * Path = the full path name where the file resides. This can be left blank if the file is in the active directory.
    * Workbook = the full workbook name, including extension 

### Macro SAVE_XLS
* Purpose: Saves the currently open spreadsheet in the same location with the same file name.
* Example Call: %save_xls;

### Macro SAVE_AS
* Purpose: Saves the currently open spreadsheet in a specified location with a specified file name.
* Example Call: %save_as(path=&top_folder\testdir, workbook=test_excel_macros_newname.xlsx);
    * Path = the full path name where we want to save the file. This can be left blank if the file is in the active directory.
    * Workbook = the full workbook name, including extension that we intend to give the newly saved file.
    * echeck = N, No, or 0 (zero) if you want to TURN OFF Excel's error checking. Generally you should leave this blank.  If left blank, Excel will pop up it's are-you-sure box when you save.  If set to N, No, or 0 the box will not pop up and Excel will behave as if you clicked that you were sure you wanted to save/overwrite.

### Macro CLOSE_XLS
* Purpose: Closes the currently open spreadsheet without saving
* Example Call: %close_xls;  
    * Note that if you have changed the spreadsheet you will get the pop-up warning from Excel when you close
* Example Call: %close_xls(echeck=N); 
    * This method will prevent the warning pop-up when closing. Use this when you know that you do not want to save any changes.

### Macro SAVE_AND_CLOSE
* Purpose: Saves and closes the currently open spreadsheet in the same location and with the same file name
* Example Call: %save_and_close;  
    * Note that if you are replacing an existing file, you may get the pop-up warning from Excel when you close
* Example Call: %save_and_close(echeck=N); 
    * This method will prevent the warning pop-up when closing. Use this when you know that you want to save your changes are are okay to overwrite any existing file.

### Macro COPYSHEET
* Purpose: Creates a copy of an existing worksheet and gives it a new name.
* Example Call: %copysheet(workbook=test_excel_macros_newname.xlsx, oldsheet=sheet1, newsheet=newsheet1, spot=1);
    * Workbook = the full workbook name, including extension
    * Oldsheet = the name of the worksheet to be copied
    * Newsheet = the name we want the newly created sheet to have.
    * Spot = the place in the worksheet where the new sheet should be placed (1 = 1st tab, 2 = 2nd, etc)

### Macro MOVESHEET
* Purpose: To move an existing worksheet to another location in the spreadsheet.
* Example Call: %movesheet(workbook=test_excel_macros_newname.xlsx, sheet2move=sheet3, spot=1);
    * Workbook = the full workbook name, including extension
    * sheet2move = the name of the worksheet to be moved
    * Spot = the place in the worksheet where the new sheet should be placed (1 = 1st tab, 2 = 2nd, etc)

### Macro RENAMESHEET
* Purpose: To rename an existing worksheet
* Example Call: %renamesheet(workbook=test_excel_macros_newname.xlsx, oldsheet=sheet2, newsheet=BlahBlahBlah);
    * Workbook = the full workbook name, including extension
    * Oldsheet = the name of the worksheet to be renamed
    * Newsheet = the mew name we want the gove the sheet

### Macro DELSHEET
* Purpose: To delete an existing worksheet
* Example Call: %delsheet(workbook=test_excel_macros_newname.xlsx, sheet2del=newsheet1, echeck=N);
    * Workbook = the full workbook name, including extension
    * sheet2del = the name of the worksheet to be deleted
    * Echeck = N, No, or 0 (zero) if you want to TURN OFF Excel’s error checking.  For this macro we will usually need this option.

### Macro SHEETINFO
* Purpose: Creates a dataset with information about the structure of the given spreadsheet
* Example Call: %sheetinfo(workbook=test_excel_macros_newname.xlsx, sheetds=sheet_info);
    * Workbook = the full workbook name, including extension
    * sheetds = the name of the SAS dataset that holds information about the spreadsheet
* How does it work?
    * Creates a new tab (Macro1) in the spreadsheet.
    * Outputs excel macro code to this new tab, runs it, then reads in the results.
    * Creates a dataset with the name you specify that contains one observation per worksheet or tab in the spreadsheet.
* What does the dataset look like?

|sheet_name      | nrows               | ncols               |
|---------------| -------------------- | --------------------|
Sheet1 | 6 | 3 
New | 11 | 3
Sheet2 | 21 | 5
Sheet3 | 13 | 5
* Important notes:
    * The macro generates a new tab, Macro1.  If there is already a Macro1 tab or if the macro is used more than once without closing and re-opening the spreadsheet then it will not work correctly.
    * The macro will overwrite anything in an already-existing Macro1 tab.
    * The SHEETINFO macro will delete the Macro1 tab when it is finished.
    * To prevent any chance of re-creating a Macro1 tab, close and reopen the spreadsheet just after running this macro.  This will not be an issue unless you intend to run the SHEETINFO macro multiple times or if you run another macro that automatically generates a macro tab.


### Macro READ_ALL_SHEETS
* Purpose: Uses the dataset created by the sheetinfo macro to read each worksheet/tab into its own dataset.
* Please note that this is super-clunky and may not give you the best output.  Consider other alternatives for reading in the data, including Python.
* Example Call: %read_all_sheets(workbook=test_excel_macros_newname.xlsx, sheetds=sheet_info,nm2use=with_excel_formats);
    * Workbook = the full workbook name, including extension
    * sheetds = the name of the SAS dataset that holds information about the spreadsheet
    * nm2use = the generic name to use for the datasets created.  Leave blank to use the tab name as the dataset name.
* How it works:
    * This macro actually creates 2 datasets for each worksheet-- an initial and a final.
    * Use macro variable nm2use if you do NOT want to use the worksheet/tab name as the dataset name
    * nm2use=data will generate datasets data1, data2, etc
    * The datasets will be numbered by the order they appear in the spreadsheet
    * Blank nm2use will name the datasets after the worksheet/tab. This will cause errors if there are spaces or special characters in the tab name.
    * Note that these datasets will be empty for sheets with 0 rows and columns.
* Dataset initial_sheet1 
    * will have as many observations as sheet1 has rows and as many variables as sheet1 has columns.
    * All data are read in as character variables with length 300.
    * Variable names are v1, v2, etc.
    * Use this dataset if you want to rename and format variables yourself.
    * All variables have format $300. and the dataset has the label “Initial Dataset from Spreadsheet test_excel_macros.xlsx  Worksheet name:  Sheet 1”
* Dataset final_sheet1
    * Cleaned-up version of initial_sheet1
    * The macro makes its best guess as to whether each variable is numeric or character.  
    * Character variables we be as long as the longest value found in the column.
    * Numeric variables will be in the best16 format if there are no decimal places.
    * Numeric variables will be in the best16.x format if there are decimal places.  Here x is the farthest decimal place found in the column.
    * If the first row appears to be a title row, then it will create labels for each of the variables and remove the first row from the dataset.
    * Variable names are col1, col2, etc.
    * The macro makes a best guess at variable format and labels. The dataset has the label “Final Dataset from Spreadsheet test_excel_macros.xlsx  Worksheet name:  Sheet 1”



sdfasdfasdf
### Macro OPEN_XLS_FILE
* Purpose: Open a specific, existing Excel file
* If Excel is not open, it will open Excel before opening the specified workbook.
* Example Call: %open_xls_file(path=&top_folder,workbook=test_excel_macros.xlsx);
    * Path = the full path name where the file resides. This can be left blank if the file is in the active directory.
    * Workbook = the full workbook name, including extension 


