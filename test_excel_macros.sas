****************************************************;
** Title:   test_excel_macros.sas                 **;
** Author:  Maria Cupples Hudson                  **;
** Purpose: To run the excel macros against a     **;
**          dummied up spreadsheet.               **;
** Date:    7/29/2010                             **;
****************************************************;
options mlogic mprint symbolgen;

%let top_folder = C:\Users\mhudson\marias_projects\SAS_Excel_Macros;
%include "&top_folder.\excel_macros.sas";

* Open up Excel;
%startxl;

* Open up the spreadsheet created for demonstrating/testing these macros;
%open_xls_file(path=&top_folder,
               workbook=test_excel_macros.xlsx);

* Save the spreadsheet with a new path and file name;
%save_as(path=&top_folder\testdir,
         workbook=test_excel_macros_newname.xlsx);

* Copy a tab within the spreadsheet;
%copysheet(workbook=test_excel_macros_newname.xlsx,
           oldsheet=sheet1,
           newsheet=newsheet1,
           spot=1);

* Save the spreadsheet with the same name/location;
* check the time stamp to see that it happened;
%save_xls;

* Delete the sheet we just created.;
%delsheet(workbook=test_excel_macros.xlsx,sheet2del=newsheet1);

* Create the copy again;
%copysheet(workbook=test_excel_macros_newname.xlsx,
           oldsheet=sheet1,
           newsheet=newsheet1,
           spot=1);

* Delete it again, this time blocking the warning popup;
%delsheet(workbook=test_excel_macros.xlsx,sheet2del=newsheet1,echeck=N);

* Move sheet3 to the front of the spreadsheet;
%movesheet(workbook=test_excel_macros_newname.xlsx,
           sheet2move=sheet3,
           spot=1);

* Rename sheet2 to be BlahBlahBlah;
%renamesheet(workbook=test_excel_macros.xlsx,
             oldsheet=sheet2,
             newsheet=BlahBlahBlah);

* Read spreadsheet description information into a dataset;
* Using the best practice for the SHEETINFO macro, we save and close;
* the file before re-openning and then reading in the information. ;
%save_and_close;
%open_xls_file(path=&top_folder\testdir,
               workbook=test_excel_macros_newname.xlsx);
%sheetinfo(workbook=test_excel_macros_newname.xlsx,
           sheetds=sheet_info);

* Again, as a best practice, close the spreadsheet without saving. Then;
* re-open. This is probably unnecessary but will prevent major problems;
* if you were to rerun the sheetinfo macro.;
%close_xls(echeck=N);
%open_xls_file(path=&top_folder\testdir,
               workbook=test_excel_macros_newname.xlsx);


* Use the sheet_info dataset to pull all tabs into SAS datasets;
* Leave excel formatting as is. Again, as a best practice close;
* and re-open to prevent problems.;
%read_all_sheets(workbook=test_excel_macros_newname.xlsx,
                 sheetds=sheet_info,nm2use=with_excel_formats);
%close_xls(echeck=N);
%open_xls_file(path=&top_folder\testdir,
               workbook=test_excel_macros_newname.xlsx);

* Use the sheet_info dataset to pull all tabs into SAS datasets;
* Remove excel formatting as is.  Note that closing and re-openning;
* is VERY IMPORTANT here because the macro is changing the formats;
* in excel. Unless you want these changes saved to the file, it is;
* best to close without saving.;
%read_all_sheets(workbook=test_excel_macros_newname.xlsx,
                 sheetds=sheet_info,nm2use=formats_removed,rmfmt=Y);
%close_xls(echeck=N);
%open_xls_file(path=&top_folder\testdir,
               workbook=test_excel_macros_newname.xlsx);



* Run proc contents on the sheet1 datasets;
proc contents data=initial_with_excel_formats2;
run;
proc contents data=final_with_excel_formats2;
run;

* Compare initial and final datasets with and without formatting removed;
title "Comparison of the initial datasets created either with Excel formatting or without";
proc compare data=initial_with_excel_formats2
             compare=initial_formats_removed2;
run;
title;

title "Comparison of the final datasets created either with Excel formatting or without";
proc compare data=final_with_excel_formats2
             compare=final_formats_removed2;
run;
title;

* Close Excel without saving;
%close_xls(echeck=N);

* Create a new spreadsheet (just to see how the macro works);
%new_xls;

* Close Excel without saving;
%close_xls(echeck=N);
