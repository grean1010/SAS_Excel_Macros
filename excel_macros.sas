****************************************************;
** Title:   excel_macros.sas                      **;
** Author:  Maria Cupples Hudson                  **;
** Purpose: To create a series of macros to help  **;
**          manipulate spreadsheets in Excel.     **;
** Date:    5/12/2010                             **;
****************************************************;

* This macro opens up the Excel Program and checks to see;
* that it opened correctly. There are no input variables.;
%macro startxl;

  filename cmds dde 'excel|system';

  options noxwait noxsync;

  data _null_;

    length fid rc start stop time 8;

    * The fopen function will return a positive integer;
    * if the filefef is working (i.e. if Excel system is open);
    fid=fopen('cmds','s');

    * If Excel is not currently open, then we run this loop;
    if (fid le 0) then do;

      * Enter a system command to open up Excel;
      rc=system('start excel');

      * record the current time in the variable start;
      start=datetime();

      * Set a final stop time.  Here we choose 100. This will;
      * allow up to 100 seconds for Excel to open.;
      stop=start+100;

      * keep poking the file reference until either the program;
      * opens or we time out;
      do while (fid le 0);
     
        * poke at it;
        fid=fopen('cmds','s');

        * reset the current time;
        time=datetime();

        * If we are past the stop time we initially set up, then;
        * set the fid to be 1 so that we can stop the loop.  Also;
        * put a note in the lst file that Excel did not open properly;
        if (time ge stop) then do;
          fid=1;
          file print;
          put "PROBLEM-- Excel did not open properly.";
        end;

      end;

    end;

    rc=fclose(fid);

  run;


%mend startxl;

* This macro creates a new/blank spreadsheet with one blank worksheet;
%macro new_xls;

  filename cmds dde 'excel|system';

  * The fopen function will return a positive integer;
  * if the filefef is working (i.e. if Excel system is open);
  data _null_;
    fid=fopen('cmds','s');
    if fid > 0 then startup = 0;
    else startup = 1;
    call symput('startup',left(trim(startup)));
    rc=fclose(fid);
  run;

  * If excel is not already open, then start it up. It will;
  * automatically create a blank spreadsheet with 3 tabs;
  %if &startup = 1 %then %do;

    %startxl;

    * Delete the extra two tabs so that this will look the;
    * same as the one-sheet created by the new command;
    data _null_;
      file cmds;
      put '[error(false)]';
      put '[workbook.delete("Sheet2")]';
      put '[workbook.delete("Sheet3")]';
      put '[error(true)]';
    run;

  %end;
  
  %else %do;

    data _null_;
      file cmds;

      * Creates a new spreadsheet with one blank sheet;
      put '[new(1)]';

    run;

  %end;

%mend new_xls;

* This macro opens up a specific Excel file. It first looks to see  ;
* if the excel system is open.  If not, then it will open excel and ;
* then the file.  Otherwise it simply opens the file in the existing;
* excel window.;
* INPUT VARIABLES:                                                  ;
* path = the full path name of the file to be opened.  This can be  ;
*        left blank to open a file in the active directory.         ;
* workbook = The full file name, including extension, of the        ;
*            spreadsheet to be opened.                              ;
%macro open_xls_file(path=,workbook=);

  data _null_;

     single = "'";
     double = '"';

     length stmt path workbook f2open $300.;

     * use the path and file name as given;
     path = "&path";
     path = left(trim(path));
     workbook = "&workbook";
     workbook = left(trim(workbook));

     * if the path was not specified then we are opening a file in the;
     * current/active directory and do no need the path name as part of the;
     * full file name;
     if path = '' then f2open = left(trim(workbook));

     * If the path name is specified then look to see if the final slash was;
     * included.  If not, add it.  Then create the full file name using both;
     * the path and file name.;
     else if substr(path,length(path),1) ne "/" then do;
       path = left(trim(path))||"/";
       f2open = left(trim(path))||left(trim(workbook));
     end;

     * Create a put statement that will open the spreadsheet;
     stmt = "put "||single||"[open("||double||left(trim(f2open))||double||")]"||single;

     * put that statement into a macro variable so we can call;
     * it in a later data _null_;
     call symput('stmt',left(trim(stmt)));

  run;

  * open Excel, test to see if the program openned correctly;
  %startxl;

  * open the excel file;
  data _null_;
    file cmds;
    * Use the statement created above to open up the file we want;
    &stmt;
  run;

%mend open_xls_file;

* This macro saves the spreadsheet under its current name and;
* location. There are no input macro variables.;
%macro save_xls;

  filename cmds dde 'excel|system';

  data _null_;
    file cmds ;
    put '[save()]';
  run;     

%mend save_xls;

* This macro saves the current excel file under a different path ;
* and file name.;
* INPUT VARIABLES:                                                  ;
* path = the full path name where we intent to save the file. This  ;
*        can be left blank to open a file in the active directory.  ;
* workbook = The full file name, including extension, that we intend;
*            to give the newly saved file.                          ;
* echeck = N, No, or 0 if you want to turn OFF error checking in    ;
*          Excel. Be careful in selecting this option. Generally you;
*          will leave echeck blank so that Excel can perform its    ;
*          checks. For this macro, that means checking that the new ;
*          file name/location does not overwrite an existing one.   ;
*          Only turn off the checks if you know for certain that you;
*          want an existing workbook overwritten.                   ;
%macro save_as(path=,workbook=,echeck=);

  data _null_;

     single = "'";
     double = '"';

     length stmt path workbook f2open $300.;

     * use the path and file name as given;
     path = "&path";
     path = left(trim(path));
     workbook = "&workbook";
     workbook = left(trim(workbook));

     * if the path was not specified then we are opening a file in the;
     * current/active directory and do no need the path name as part of the;
     * full file name;
     if path = '' then f2open = left(trim(workbook));

     * If the path name is specified then look to see if the final slash was;
     * included.  If not, add it.  Then create the full file name using both;
     * the path and file name.;
     else if substr(path,length(path),1) ne "/" then do;
       path = left(trim(path))||"/";
       f2open = left(trim(path))||left(trim(workbook));
     end;

     * Create a put statement that will open the spreadsheet;
     stmt = "put "||single||"[save.as("||double||left(trim(f2open))||double||")]"||single;

     call symput('stmt',left(trim(stmt)));

  run;

  filename cmds dde 'excel|system';

  * save the file with the new path and file name;
  data _null_;
    file cmds;

    * Look at the echeck macro variable;
    echeck = "&echeck";
    echeck = substr(left(trim(echeck)),1,1);
    echeck = upcase(echeck);

    * If we do not want Excel to pop up a warning box, then turn;
    * off that capability;
    if echeck in ("N","0") then put '[error(false)]';

    &stmt;

    * Make sure you turn error-checking back on.;
    put '[error(true)]';

  run;

%mend save_as;

* This macro closes the Excel System.  The macro variable echeck;
* should be N, No, or 0 if you want to turn OFF error checking;
* in excel. Be careful in selecting this option. Generally you;
* will leave echeck blank so that Excel will pop up a warning;
* message if need be. In this case turning off the error-checks;
* will allow you to close without saving. While there may be some;
* cases where this is desirable, you will usually not use this;
* option;
%macro close_xls(echeck=);

  filename cmds dde 'excel|system';

  * save the file with the new path and file name;
  data _null_;
    file cmds;

    * Look at the echeck macro variable;
    echeck = "&echeck";
    echeck = substr(left(trim(echeck)),1,1);
    echeck = upcase(echeck);

    * If we do not want Excel to pop up a warning box, then turn;
    * off that capability;
    if echeck in ("N","0") then put '[error(false)]';

    put '[quit()]';

  run;

%mend close_xls;

* This macro saves the current file under the current name and;
* then closes the Excel System.  The macro variable echeck;
* should be N, No, or 0 if you want to turn OFF error checking;
* in excel. Be careful in selecting this option. Generally you;
* will leave echeck blank so that Excel will pop up a warning;
* message if need be.;
%macro save_and_close(echeck=);

  filename cmds dde 'excel|system';

  data _null_;
    file cmds ;

    * Look at the echeck macro variable;
    echeck = "&echeck";
    echeck = substr(left(trim(echeck)),1,1);
    echeck = upcase(echeck);

    * If we do not want Excel to pop up a warning box, then turn;
    * off that capability;
    if echeck in ("N","0") then put '[error(false)]';

    put '[save()]';
    put '[quit()]';

  run;     

%mend save_and_close;

* This macro will copy one sheet of a workbook, place it where specified;
* in the spreadsheet, and rename it. Input variables are defined as follows;
* workbook = The name of the excel file WITH EXTENSION (ex. book1.xlsx);
* oldsheet = The name of the worksheet to be copied (ex. sheet1);
* newsheet = The name of the sheet to be created (ex. newsheet1);
* spot = The place in the spreadsheet where you want the new worksheet;
*        placed.  spot = 1 means the new sheet will be the first one.;
%macro copysheet(workbook=,oldsheet=,newsheet=,spot=);

  filename cmds dde 'excel|system';

  data _null_;

    length copy_command rename_command $100.;

    single = "'";
    double = '"';


    copy_command = "put '[workbook.copy("||double||"&oldsheet"||double||","||double||"&workbook."||double||",&spot.)]'";
    rename_command = "put '[workbook.name("||double||"&oldsheet. (2)"||double||","||double||"&newsheet"||double||")]'";
  
    call symput('copy_command',left(trim(copy_command)));
    call symput('rename_command',left(trim(rename_command)));

  run;

  %put COPY COMMAND USED:  &copy_command;
  %put RENAME COMMAND USED:  &rename_command;

  data _null_;
    file cmds ;
    &copy_command;
    &rename_command;
  run;

%mend copysheet;


* This macro will move one sheet of a workbook and place it where specified;
* in the spreadsheet.  Input variables are defined as follows;
* workbook = The name of the excel file WITH EXTENSION (ex. book1.xlsx);
* sheet2move = The name of the worksheet to be moved (ex. sheet1);
* spot = The place in the spreadsheet where you want the new worksheet;
*        placed.  spot = 1 means the new sheet will be the first one.;
%macro movesheet(workbook=,sheet2move=,spot=);

  filename cmds dde 'excel|system';

  data _null_;

    length move_command $100.;

    single = "'";
    double = '"';

    move_command = "put '[workbook.move("||double||"&sheet2move."||double||","||double||"&workbook."||double||",&spot.)]'";
  
    call symput('move_command',left(trim(move_command)));

  run;

  %put MOVE COMMAND USED:  &move_command;

  data _null_;
    file cmds ;
    &move_command;
  run;

%mend movesheet;

* This macro will rename one sheet of a workbook;
* Input variables are defined as follows;
* workbook = The name of the excel file WITH EXTENSION (ex. book1.xlsx);
* oldsheet = The name of the worksheet to be renamed (ex. sheet2);
* newsheet = The new name of the sheet (ex. newname);
%macro renamesheet(workbook=,oldsheet=,newsheet=);

  filename cmds dde 'excel|system';

  data _null_;

    length rename_command $100.;

    single = "'";
    double = '"';

    rename_command = "put '[workbook.name("||double||"&oldsheet."||double||","||double||"&newsheet"||double||")]'";
  
    call symput('rename_command',left(trim(rename_command)));
  run;

  %put RENAME COMMAND USED:  &rename_command;

  data _null_;
    file cmds ;
    &rename_command;
  run;

%mend renamesheet;

* This macro will delete a sheet in a workbook.;
* Input variables are defined as follows;
* workbook = The name of the excel file WITH EXTENSION (ex. book1.xlsx);
* sheet2del = The name of the worksheet to be deleted (ex. sheet2);
* echeck = N, No, or 0 if you want to turn OFF error checking in    ;
*          Excel. Be careful in selecting this option. For this     ;
*          macro (unlike others) you will usually want to turn off  ;
*          the error checks.  Otherwise a warning will pop up each  ;
*          time you run and your program will hang until you click  ;
*          the OK box.                                              ;
%macro delsheet(workbook=,sheet2del=,echeck=);

  filename cmds dde 'excel|system';

  data _null_;

    length delete_command $100.;

    single = "'";
    double = '"';

    delete_command = "put '[workbook.delete("||double||"&sheet2del."||double||")]'";
  
    call symput('delete_command',left(trim(delete_command)));

  run;

  %put DELETE COMMAND USED:  &delete_command;

  * Run the delete command.  Note that we need to turn error-detection off;
  * in excel to prevent a box from popping up and asking you to click OK;
  data _null_;
    file cmds ;

    * Look at the echeck macro variable;
    echeck = "&echeck";
    echeck = substr(left(trim(echeck)),1,1);
    echeck = upcase(echeck);

    * If we do not want Excel to pop up a warning box, then turn;
    * off that capability;
    if echeck in ("N","0") then put '[error(false)]';

    * Send the delete command;
    &delete_command;

    * Turn Excel error detection back on;;
    put '[error(true)]';

  run;

%mend delsheet;

* This macro creates a new worksheet in an existing workbook. Because;
* the naming is automatic (Sheet1, Sheet2, etc), we cannot be certain;
* of the name the new sheet will have.  It will create whatever is next;
* in line based on what is in the spreadsheet.  In order to have full;
* control over the name, delete Sheet1 from the workbook then add a new;
* sheet and rename it.;
* Input variables are defined as follows;
* workbook = The name of the excel file WITH EXTENSION (ex. book1.xlsx);
%macro new_worksheet(workbook=);

  * open the excel file and insert macro sheet;
  filename cmds dde 'excel|system' lrecl=200;

  data _null_;

    file cmds;

    * Move to the next available spot in the workbook;
    * This is a precautionary step so that we do not create;
    * mulitple worksheets if multiple sheets happen to be;
    * selected when this macro is run;
    put '[workbook.next()]';

    * Create a blank worksheet;
    put '[workbook.insert(3)]';

  run;

%mend new_worksheet;

* Macro Sheetinfo;
* This macro will look at an excel spreadsheet and read information from it;
* including the number of worksheets, the worksheet name, and the number of;
* rows and columns in each worksheet.;
* CAUTIONS:
* 1) If a worksheet with the name "Macro1" already exists in the workbook, ;
*    the sheetinfo macro will not run correctly. ;
* 2) Best practice for running this macro is to save and close the spreadsheet.;
*    Re-open it. Run the sheetinfo macro. Close without saving. Then re-open.;
* Input variables are defined as follows;
* workbook = The name of the excel file WITH EXTENSION (ex. book1.xlsx);
* sheetds = the name of the dataset that we will be creating to contain;
*           information abotu this workbook.;
%macro sheetinfo(workbook=,sheetds=);

  * open the excel file and insert macro sheet;
  filename cmds dde 'excel|system' lrecl=200;

  data _null_;

    file cmds;

    * Move to the next available spot in the workbook;
    put '[workbook.next()]';

    * Create a blank worksheet;
    put '[workbook.insert(3)]';

  run;

  * Move the macro1 worksheet to the front of the file;
  %movesheet(workbook=&workbook,sheet2move=Macro1,spot=1);

  * Set up a file reference for the Macro1 worksheet that we just created;
  filename xlmacro dde "excel|macro1!r1c1:r1000c1" notab lrecl=200;

  * Initialize the number sheets to be blank;
  %let nsheets=;

  data _null_;
    file xlmacro;

    * We will create an excel macro statement in the cell A1;
    * The output from that excel macro will be put into the selcted;
    * cell (in this case row 1 column 2 or A2);
    put '=select("r1c2:r1c2")';

    * The following statements create an excel macro that finds;
    * the number of sheets in the workbook;
    put '=set.name("nsheet",selection())';
    put '=set.value(nsheet,get.workbook(4))';
    put '=halt(true)';
    put '!dde_flush';

  run;

  data _null_;

    * These statements run the macro and prevent excel from;
    * flashing an error box.;
    file cmds;
    put '[run("macro1!r1c1")]';
    put '[error(false)]';

  run;

  * Create a file reference to Cell A2 that we just populated;
  filename nsheets dde "excel|macro1!r1c2:r1c2" notab lrecl=200;

  * Read in the number of spreadsheets from that cell;
  data _null_;
    length nsheets 8;
    infile nsheets;
    input nsheets;
    call symput('nsheets',trim(left(put(nsheets,2.))));
  run;

  * Because we created the sheet Macro1 for the sole purpose of;
  * finding the number of sheets, we need to subtract 1 from the;
  * total number of sheets.;
  %let nsheets=%eval(&nsheets-1);

  * Clear out teh file name;
  filename nsheets clear;

  * Clear the information we wrote to the Macro1 worksheet;
  * Then activate the Macro1 worksheet. We will use this as;
  * our starting point as we loop through each sheet.;
  data _null_;
    file cmds;
    put '[workbook.activate("macro1")]';
    put '[select("r1c1:r1000c2")]';
    put '[clear(1)]';
    put '[select("r1c1")]';
  run;

  %put Number of Sheets = &nsheets;

  * Create a file reference to the first 1000 rows and 100 columns of the macro1 spreadsheet;
  filename m1sheet dde "excel|macro1!r1c1:r1000c100" notab lrecl=200;

  data _null_;

    file m1sheet;
    length maccmd $200.;

    * Find the name of the ith sheet;
    %do i=1 %to &nsheets;

      maccmd="=select(!$b$&i,!$b$&i)";
      put maccmd;
      put '=set.name("cell",selection())';

      %do k=1 %to &i;
        put '=workbook.next()';
      %end;

      put '=set.value(cell,get.workbook(3))';
      put '=workbook.activate("Macro1")';

    %end;

    * Find the number of rows in the ith sheet;
    %do i=1 %to &nsheets;

      maccmd="=select(!$c$&i,!$c$&i)";
      put maccmd;
      put '=set.name("rows",selection())';

      %do k=1 %to &i;
        put '=workbook.next()';
      %end;

      put '=set.value(rows,get.document(10))';
      put '=workbook.activate("Macro1")';

    %end;

    * Find the number of columns in the ith sheet;
    %do i=1 %to &nsheets;

      maccmd="=select(!$d$&i,!$d$&i)";
      put maccmd;
      put '=set.name("cols",selection())';

      %do k=1 %to &i;
        put '=workbook.next()';
      %end;

      put '=set.value(cols,get.document(12))';
      put '=workbook.activate("Macro1")';

    %end;

    put '=halt(true)';
    put '!dde_flush';

    * Now run the macro we just created;
    file cmds;
    put '[run("macro1!r1c1")]';
    put '[error(false)]';

  run;

  * Read in the results of the macro run;
  data &sheetds;
    length sheet_name v1-v4 $200. nrows ncols 8.;
    infile m1sheet dsd delimiter='09'x notab TRUNCOVER;

    input v1 $ v2 $ v3 $ v4 $;

    sheet_name =reverse(scan(reverse(v2),1,']'));

    nrows = input(v3,8.);
    ncols = input(v4,8.);

    if nrows = . then nrows = 0;
    if ncols = . then ncols = 0;

    keep sheet_name nrows ncols;
    if sheet_name ne "";

  run;

  title "Information For the &nsheets Worksheets in the workbook &workbook";
  proc print data=&sheetds noobs;
    var sheet_name nrows ncols;
  run;
  title;

  * delete the macro1 sheet now that we are done with it;
  %delsheet(workbook=&workbook,sheet2del=macro1,echeck=N);

%mend sheetinfo;

* Macro Read_All_Sheets;
* This macro uses the information collected by the sheetinfo macro;
* to read all of the information from an excel spreadsheet into datasets;
* Input variables are defined as follows;
* workbook = The name of the excel file WITH EXTENSION (ex. book1.xlsx);
* sheetds = the name of the dataset created by the sheetinfo macro.;
* nm2use = the generic name to give the datasets created for each worksheet;
*          ex. if nm2use=data then datasets will be named data1, data2, data3, etc;
*          Leave nm2use blank if you want to use the tab name as the dataset name.;
* alltitles = Y, Yes, or 1 if you know the first row of the sheets are title lines;
* notitles = N, No, or 0 if you know the first row of the sheets are NOT title lines;
* rmfmt = Y, Yes, or 1 if you want to remove all excel formatting before reading;
*         in the data from the spreadsheets;
* NOTES:  The macro actually creates two datasets for each tab/worksheet-an initial;
*         and a final dataset. Empty datasets will be created for worksheets with ;
*         no rows or columns.;
*         If alltitles and notitles are left blank then the macro will a attempt to;
*         guess if the first row is a title line or not for each worksheet. In all;
*         cases the initial dataset will contain all rows.  Only the final will delete;
*         title lines.;
%macro read_all_sheets(workbook=,sheetds=,nm2use=,alltitles=,notitles=,rmfmt=);

  data test;
    set &sheetds end=last;
    
    length row col name $15. nm2use $200.;

    nm2use = "&nm2use";
    nm2use = left(trim(upcase(nm2use)));

    * If the macro variable is blank then we use the name;
    * of the worksheet as the dataset name.;
    if nm2use = "" then nm2use = sheet_name;

    * Otherwise just use whatever generic name specified.;
    else if nm2use ne '' then nm2use = left(trim(nm2use))||left(trim(_n_));

    row="row"||left(trim(_n_));
    col="col"||left(trim(_n_));
    name="name"||left(trim(_n_));
    dsname="dsname"||left(trim(_n_));

    * Store the dataset information into macro variables;
    call symput(row,compress(trim(nrows)));
    call symput(col,compress(trim(ncols)));
    call symput(name,left(trim(sheet_name)));
    call symput(dsname,left(trim(nm2use)));

    * store the number of spreadsheets to be converted into datasets;
    if last then call symput('nsheets',left(trim(_n_)));

    * Standardize the rmfmt input macro variable;
    rmfmt = "&rmfmt";
    rmfmt = left(trim(upcase(rmfmt)));
    rmfmt = substr(rmfmt,1,1);
    if rmfmt = '1' then rmfmt = "Y";
    call symput('rmfmt',left(trim(rmfmt)));

  run;

  * Create dataset label statements so that we can tie both;
  * the spreadsheet and tab name to each dataset;
  data _null_;
    file "dslabels.inc";
    
    put "proc datasets lib=work nolist;";

    %do i = 1 %to &nsheets;
  
      double = '"';
      length ilabel2apply flabel2apply $300.;
      ilabel2apply = double||"Initial Dataset from Spreadsheet &workbook.  Worksheet Name: &&&name&i"||double;
      flabel2apply = double||"Final Dataset from Spreadsheet &workbook.  Worksheet Name: &&&name&i"||double;

      put "  modify initial_&&&dsname&i (label= " ilabel2apply ");";
      put "  modify final_&&&dsname&i (label= " flabel2apply ");";

    %end;

    put "  quit;";
    put "run;";

  run;


  * Read in each worksheet;
  %do i = 1 %to &nsheets;

    %put "For Sheet number &i, we have the following information";
    %put "  Dataset name:  &&&dsname&i";
    %put "  Number of Rows:  &&&row&i";
    %put "  Number of Columns:  &&&col&i";
    %put "  Sheet Name:  &&&name&i";

    * If there are rows and columns then read them into a dataset;
    %if (%eval(&&&row&i) > 0 and %eval(&&&col&i) > 0) %then %do;

      * Set up a file reference to this worksheet;
      filename sheet&i dde "excel|&&&name&i!r1c1:r&&&row&i..c&&&col&i" notab lrecl=200;
 

      * Remove Excel formatting, if called for by the input macro variable;
      %if "&rmfmt" = "Y" %then %do;

        * First select all rows and columns in the sheet and change them;
        * to text format. This removes any and all Excel formatting. This;
        * is especially important for numeric variables and decimal places.;
        * NOTE, this data _null_ only creates the select statement.  The ;
        * next data _null_ actually runs it IF the rmfmt variable indicates;
        * that we want excel formatting removed.;
        data _null_;
  
          file cmds;
        
          single = "'";
          double = '"';

          length stmt stmt2 stmt3 $200.;
          stmt = "put "||single||"[column.width(0,"||double||"c1:c&&&col&i"||double||",false,3)]"||single||";";
          call symput('stmt',left(trim(stmt))); 

          stmt2 = "put "||single||"[select("||double||"r1c1:r&&&row&i..c&&&col&i"||double||")]"||single||";";
          call symput('stmt2',left(trim(stmt2))); 

          stmt3 = "put "||single||"[workbook.select("||double||"&&&name&i"||double||")]"||single||";";
          call symput('stmt3',left(trim(stmt3))); 

        run;  

        data _null_;
          file cmds;

          * select the worksheet we are currently reading in;
          &stmt3;

          * select only the cells we will be reading in;
          &stmt2;

          * Format the numbers to be plain text;
          put '[format.number("@")]'; 
       

          * make sure the columns are in the best-fit format;
          * to prevent accidental truncation or rounding;
          &stmt;

        run;

      %end;
        
      data initial_&&&dsname&i;
  
        infile sheet&i dsd delimiter='09'x notab;

        * Read everything in as character;
        length v1 - v&&&col&i $300.;

        input %do k = 1 %to &&&col&i;

                v&k $
     
              %end;

              ;

      run;

      * Test to see if row 1 is a title line;
      * Also test to see if each observation is character or numeric;
      data test_&&&dsname&i;
        set initial_&&&dsname&i end=last;

        * Set lengths for titles and variable type indicators;
        length tt1 - tt&&&col&i $200. 
               nums1 - nums&&&col&i 
               dens1 - dens&&&col&i 
               pctnum1 - pctnum&&&col&i
               wnums1 - wnums&&&col&i 
               wdens1 - wdens&&&col&i 
               wpctnum1 - wpctnum&&&col&i 
               maxlen1 - maxlen&&&col&i 8.
               rpt1 - rpt&&&col&i 3.
               numchar1 - numchar&&&col&i $1. ;

        * array for the text read in from the spreadsheet;
        array vars(*) v1 - v&&&col&i;
 
        * array for titles. These hold the value in the first observation;
        * so we can look to see if that appears to be a title line;
        array ttl(*) $ tt1 - tt&&&col&i;

        * Indicators for whether or not the value of the first row is repeated;
        array rpt(*) rpt1 - rpt&&&col&i;

        * Numerators/denominators/percents for all but 1st obs;
        array nums(*) nums1 - nums&&&col&i;
        array dens(*) dens1 - dens&&&col&i;
        array pcts(*) pctnum1 - pctnum&&&col&i;

        * Numerators/denominators/percents for all obs;
        array wnums(*) wnums1 - wnums&&&col&i;
        array wdens(*) wdens1 - wdens&&&col&i;
        array wpcts(*) wpctnum1 - wpctnum&&&col&i;

        * the maximum length the variable takes;
        array maxlen(*) maxlen1 - maxlen&&&col&i;

        * the maximum number of decimal places the variable has;
        array maxdec(*) maxdec1 - maxdec&&&col&i;
 
        * final numeric/character flag for variables;
        array numchar(*) $ numchar1 - numchar&&&col&i;

        if _n_ = 1 then do i = 1 to dim(ttl);

          * initialize numerators and denominators to zero;
          nums(i) = 0;
          dens(i) = 0;
          wnums(i) = 0;
          wdens(i) = 0;

          * initialize the maximum length and decimal place to zero;
          maxlen(i) = 0;
          maxdec(i) = 0;

          * set the title variables equal to the 1st value of each variable;
          ttl(i) = vars(i);

          * set repeat flags to zero;
          rpt(i) = 0;

        end;

        retain tt1 - tt&&&col&i nums1 - nums&&&col&i dens1 - dens&&&col&i
               wnums1 - wnums&&&col&i wdens1 - wdens&&&col&i 
               maxlen1 - maxlen&&&col&i maxdec1 - maxdec&&&col&i
               rpt1 - rpt&&&col&i;

        do i = 1 to dim(ttl);

          if vars(i) ne '' then do;

            * Count the number of non-blank observations for each variable;
            wdens(i) = wdens(i) + 1;

            * Count the number of non-blank observations for each variable;
            * DISREGARDING the first observation;
            if _n_ ne 1 then dens(i) = dens(i) + 1;

            * Count the number of numeric observations for each variable;
            if input(vars(i),8.) ne . then wnums(i) = wnums(i) + 1;

            * Count the number of numeric observations for each variable;
            * DISREGARDING the first observation;
            if _n_ ne 1 and input(vars(i),8.) ne . then nums(i) = nums(i) + 1;

            * if the length of this variable is longer than the previously;
            * stored maximum, then replace it;
            if maxlen(i) < length(vars(i)) then maxlen(i) = length(vars(i));

            * If there is a decimal in the variable, then we find the number;
            * of places by subtracting the place where the decimal appears;
            * from the variable length.  If that number is greater than the;
            * previously stored maximum decimal place, then we replace it.;
            
            if index(vars(i),'.') > 0 then  
               maxdec(i) = max(maxdec(i),length(vars(i)) - index(vars(i),'.'));

            * If this is not the first observation then check to see if it;
            * is a repeat of the title line. Generally a title is different;
            * from the values within the variable.;
            if _n_ > 1 and vars(i) = ttl(i) then rpt(i) = 1;

          end;

        end;

        if last then do;

          * Pull in the alltitle and notitle flags;
          alltitles = "&alltitles";
          notitles = "&notitles";

          alltitles = substr(left(trim(upcase(alltitles))),1,1);
          notitles = substr(left(trim(upcase(notitles))),1,1);


          * initialize the missing title flag to be zero;
          misstitle = 0;

          * initialize the number of numeric titles to zero;
          numttls = 0; 

          * initialize the number of repeated titles to zero;
          rptcount = 0;

          do i = 1 to dim(ttl);

            * initialize character-numeric flag to be character;
            numchar(i) = "C";

            * Find the percentage of observations that are numeric;
            * both with and disregarding the first observation;
            pcts(i) = 100 * nums(i) / dens(i);
            wpcts(i) = 100 * wnums(i) / wdens(i);

            * Look to see if all of the titles are non-missing;
            * The first row can only be a title row if all nonmissing;
            if ttl(i) = '' then misstitle = 1;

            * Count the number of nonmissing, numeric titles;
            if ttl(i) ne '' and input(ttl(i),8.) ne . then numttls = numttls + 1;

            * Count the number of titles that are repeated in later obs;
            rptcount = rptcount + rpt(i);

          end;

          * Create the final percentage of titles that are numeric.;
          * Note that we only do this if there are no missing titles;
          if misstitle = 0 then pctnumttls = 100 * numttls / &&&col&i;

        end;

        * Create a final flag for whether or not we have a title row;
        * Start with the assumption that the first row is not a title line;
        title_line = 0;

        * The following situations indicate that the first row is;
        * not a title line:                                      ;
        * 1.  There are missing titles.                          ;
        * 2.  There are numeric titles.                          ;
        * 3.  The titles values are repeated later .             ;
        * 4.  The macro variable notitles was set to 1 or yes.   ;
        if misstitle = 0 and rptcount = 0 and pctnumttls = 0 then title_line = 1;

        * If we know all first rows are title lines, then reset the flag to 1;
        if alltitles in ('Y','1') then title_line = 1;

        * If we know we have no title lines, then reset the flag to zero;
        if notitles in ('Y','1') then title_line = 0;
 
        * Create a macro variable with the misstitle value;
        call symput('title_line',left(trim(title_line)));


        * Setting up final numeric/character flags;
        do i = 1 to dim(ttl);
            
          if title_line = 0 then do;

            * If all observations are numeric, then this is probably a;
            * numeric variable. Reset the indicator to N for this variable;
            if wpcts(i) = 100 then numchar(i) = "N";

          end;

          * If we believe that the first line has titles, then determine if;
          * all of the observations (disregarding the first) are numeric;
          * If so, then this is probably a numeric variable.  Reset the;
          * indicator to N;
          else do;
            if pcts(i) = 100 then numchar(i) = "N";
          end;


        end;

        * Keep only the last record with the summary/variable information;
        if last then output;
        keep tt1 - tt&&&col&i pctnum1 - pctnum&&&col&i wpctnum1 - wpctnum&&&col&i
             numchar1 - numchar&&&col&i title_line maxlen1 - maxlen&&&col&i
             maxdec1 - maxdec&&&col&i misstitle rptcount notitles alltitles 
             pctnumttls title_line rpt1 - rpt&&&col&i;

      run;

      * Create include files that will clean up the initial dataset;
      data _null_;
        set test_&&&dsname&i;

        single = "'";

        array ttl(*) $ tt1 - tt&&&col&i;

        * If the first row was a title row, then we create an include;
        * file with label statements.;    
        if title_line = 1 then do;

          * Add single quotes around the titles so we can use these;
          * in the label statements;
          do i = 1 to dim(ttl);
            ttl(i) = "'"||left(trim(ttl(i)))||"'";
          end;

          file "labels.inc";
          put "label ";


          %do k = 1 %to &&&col&i;
             
            put "  col&k = " tt&k ;

          %end;

          put ";";

        end;

        * Now we need to format each variable.  Use the character/numeric;
        * indicator as well as the maximum length of each column.;

        array maxlen(*) maxlen1 - maxlen&&&col&i;
        array maxdec(*) maxdec1 - maxdec&&&col&i;
        array numchar(*) $ numchar1 - numchar&&&col&i;

        length lenst1 - lenst&&&col&i $15.;
        array lenst(*) $ lenst1 - lenst&&&col&i;

        length lenst2_1 - lenst2_&&&col&i $15.;
        array lenst_2(*) $ lenst2_1 - lenst2_&&&col&i;

        do i = 1 to &&&col&i;

          * reset 0 length to be 1 since that is the minimum character length;
          if numchar(i) = "C" and maxlen(i) = 0 then maxlen(i) = 1;

          * if the variable is numeric, then we want a numeric format with;
          * plenty of room.  Make sure to include enough decimal places by;
          * using the max decimal place;
          if numchar(i) = "N" then do;
           
            * Use length 8 for numeric variables unless the maximum length;
            * is longer than that.;
            if maxlen(i) > 8 then temp = maxlen(i);
            else temp = 8;

            * the first length statement is for values with no decimal place;
            lenst(i) = left(trim(temp))||".";

            * The second length statement is for values with decimal places;
            lenst_2(i) = left(trim(temp))||"."||left(trim(maxdec(i)));

          end;

          * if it was character, then we allow for enough spaces to;
          * accomodate the longest observation.;
          else if numchar(i) = "C" then lenst(i) = "$"||left(trim(maxlen(i)))||".";

          lenst(i) = compress(lenst(i));
          lenst_2(i) = compress(lenst_2(i));

        end;          

        file "lengths.inc";

        put "format ";
        
        %do k = 1 %to &&&col&i;
            
          if (numchar&k = "C" or maxdec&k = 0) then put "  col&k " lenst&k ;
          else put "  col&k " lenst2_&k;

        %end;

        put ";";

        * Now create statements that will populate the final variables;
        file "createvars.inc";

        %do k = 1 %to &&&col&i;

          length tempstmt tempstmt2 $400.;

          if numchar&k = "C" then do;
            tempstmt="col&k = v&k;" ;
            tempstmt2="";
          end;
          else do;
            tempstmt = "if index(v&k,'.')=0 then col&k = input(left(trim(v&k)),"||compress(lenst&k)||");";
            tempstmt2 = "else col&k = input(left(trim(v&k)),"||compress(lenst2_&k)||");";
          end;

          put tempstmt;
          put tempstmt2;
          put ;

          tempstmt="";
          tempstmt2="";

        %end;

        * Finally, create a keep statements;
        file "keepers.inc";
        
        put "keep ";

        %do k = 1 %to &&&col&i;

           put "  col&k";
        
        %end;

        put ";";

      run;       

      options noxwait noxsync;

      data final_&&&dsname&i;
        set initial_&&&dsname&i;

        * Create length statements for the variables to be created;
        %include "lengths.inc";

        * If we do have a title row, then include the label statements;
        * and exclude the first observation from the final dataset;
        %if &title_line = 1 %then %do;

          %include "labels.inc";
          if _n_ = 1 then delete;

        %end;

        * Populate the variables;
        %include "createvars.inc";

        * Include the keep statements;
        %include "keepers.inc";

      run;

      * delete the include files;
      x "del lengths.inc";
      x "del labels.inc";
      x "del createvars.inc";
      x "del keepers.inc";

      * delete the temporary dataset;
      proc datasets lib=work;
        delete test_&&&dsname&i;
        quit;
      run;

    %end;

    * If there no rows or columns then create a blank dataset;
    %if %eval(&&&row&i) = 0 or %eval(&&&col&i) = 0 %then %do;

      data initial_&&&dsname&i final_&&&dsname&i;
      run;

    %end;

  %end;

  * Include the file that labels all of the datasets we just created;
  %include "dslabels.inc";

  * delete the include file now that we no longer need it;
  x "del dslabels.inc";

  * delete the temporary dataset;
  proc datasets lib=work;
    delete test;
    quit;
  run;

 %mend read_all_sheets;


