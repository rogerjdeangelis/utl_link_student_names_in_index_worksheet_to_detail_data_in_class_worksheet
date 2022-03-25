# utl_link_student_names_in_index_worksheet_to_detail_data_in_class_worksheet
Link student names in index worksheet to detail data in class worksheet    Hyperlink to the excel specific cells.    Two Solutions.     1. Libname.     2. ods excel and proc report.     Just output hyperlinks using.     name=cats('=HYPERLINK("[class.xlsx]class!A',put(seq,2.),'","',name,'")');
    Link student names in index worksheet to detail data in class worksheet

    output excel workbook
    https://goo.gl/SNBjba
    https://github.com/rogerjdeangelis/utl_link_student_names_in_index_worksheet_to_detail_data_in_class_worksheet/blob/master/utl_link_student_names_in_index_worksheet_to_detail_data_in_class_worksheet.

    Original topic: Hyperlink to the excel specific cells.

      Two Solutions.
        1. Libname.
        2. ods excel and proc report.

       Just output hyperlinks using.

       name=cats('=HYPERLINK("[class.xlsx]class!A',put(seq,2.),'","',name,'")');

    INPUT
    ==================================================

     SASHELP.CLASS total obs=19

       NAME       SEX    AGE    HEIGHT    WEIGHT

       Alfred      M      14     69.0      112.5
       Alice       F      13     56.5       84.0
       Barbara     F      13     65.3       98.0
       Carol       F      14     62.8      102.5
       Henry       M      14     63.5      102.5
       James       M      12     57.3       83.0


    PROCESS
    =======

      1. LIBNAME

          %utlfkil(d:/xls/class.xlsx);
          libname xel "d:/xls/class.xlsx";

          data xel.index(keep=name) xel.class(drop=seq);
           length name $171;
           set sashelp.class;
           output xel.class;
           seq=_n_+1;
           name=cats('=HYPERLINK("[class.xlsx]class!A',put(seq,2.),'","',name,'")');
           putlog name;
           output xel.index;
          run;quit;

          libname xel clear;

      2. ODS EXCEL

         ods excel file="d:/xls/odsexcel.xlsx" ;

         ods excel options(sheet_name="index");
         proc report data=index nowd;
           run;quit;

         ods excel options(sheet_name="class");
         proc print data=sashelp.class;
           run;quit;
         ods excel close;

      * you may need to activate the links just once and then save.
      * Highlight the links and replace ! with !;
      * Tis will activate the links;


    OUTPUT  (Two sheets in workbook d:/xls/class.xlsx)
    ==================================================

          +---------------------------------------------------+
          |                    A                              |
          |---------------------------------------------------+
       1  |  =HYPERLINK("[class.xlsx]class!A2","Alfred")      |   ** links to row Alfred in CLASS worksheet
          |---------------------------------------------------|
       2  |  =HYPERLINK("[class.xlsx]class!A3","Alice")       |
          |---------------------------------------------------|
       3  |  =HYPERLINK("[class.xlsx]class!A4","Barbara")     |
          |---------------------------------------------------|
       4  |  =HYPERLINK("[class.xlsx]class!A5","Carol")       |
          |---------------------------------------------------|
       5  |  =HYPERLINK("[class.xlsx]class!A6","Henry")       |
          |---------------------------------------------------|
       6  |  =HYPERLINK("[class.xlsx]class!A7","James")       |
          |---------------------------------------------------|
       7  |  =HYPERLINK("[class.xlsx]class!A8","Jane")        |
          |---------------------------------------------------|
       8  |  =HYPERLINK("[class.xlsx]class!A9","Janet")       |
          |---------------------------------------------------|
       9  |  =HYPERLINK("[class.xlsx]class!A10","Jeffrey")    |
          |---------------------------------------------------|
      10  |  =HYPERLINK("[class.xlsx]class!A11","John")       |
          |---------------------------------------------------|
      11  |  =HYPERLINK("[class.xlsx]class!A12","Joyce")      |
          +---------------------------------------------------+
    .....

    [INDEX]


          +----------------------------------------------------------------+
          |     A      |    B       |     C      |    D       |    E       |
          +----------------------------------------------------------------+
       1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |   ** focus on this row
          +------------+------------+------------+------------+------------+
       2  | ALFRED     |    M       |    15      |    69      |  112.5     |
          +------------+------------+------------+------------+------------+
       3  | BARBARA    |    F       |    13      |    58      |  101.5     |
          +------------+------------+------------+------------+------------+
           ...
          +------------+------------+------------+------------+------------+
       20 | WILLIAM    |    M       |    15      |   66.5     |  112       |
          +------------+------------+------------+------------+------------+

    [CLASS]

    see
    https://goo.gl/UXJKGe
    https://communities.sas.com/t5/ODS-and-Base-Reporting/hyperlink-to-the-excel-specific-cells-when-using-ODS-excel-xp/m-p/420207


    *                _              _       _
     _ __ ___   __ _| | _____    __| | __ _| |_ __ _
    | '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |
    | | | | | | (_| |   <  __/ | (_| | (_| | || (_| |
    |_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|

    ;

      sashelp.class
    *          _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __
    / __|/ _ \| | | | | __| |/ _ \| '_ \
    \__ \ (_) | | |_| | |_| | (_) | | | |
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|

    ;

    %utlfkil(d:/xls/class.xlsx);
    libname xel "d:/xls/class.xlsx";

    data xel.index(keep=name) xel.class(drop=seq);
      length name $171;
      set sashelp.class;
      output xel.class;
      seq=_n_+1;
      name=cats('=HYPERLINK("[class.xlsx]class!A',put(seq,2.),'","',name,'")');
      putlog name;
      output xel.index;
    run;quit;

    libname xel clear;

    *                          _
     _ __ ___ _ __   ___  _ __| |_
    | '__/ _ \ '_ \ / _ \| '__| __|
    | | |  __/ |_) | (_) | |  | |_
    |_|  \___| .__/ \___/|_|   \__|
             |_|
    ;

    data index(keep=name) class(drop=seq);
      length name $96;
      set sashelp.class;
      output class;
      seq=_n_+1;
      name=cats('=HYPERLINK("[class.xlsx]class!A',put(seq,2.),'","',name,'")');
      putlog name;
      output index;
    run;quit;

    %utlfkil(d:/xls/odsexcel.xlsx);
    ods excel file="d:/xls/odsexcel.xlsx" ;
    ods excel options(sheet_name="index");
    proc report data=index nowd;
    run;quit;
    ods excel options(sheet_name="class");
    proc print data=sashelp.class;
    run;quit;
    ods excel close;
    
    _ __   _____      _____ _ __    _____  _____ ___| |         
   | `_ \ / _ \ \ /\ / / _ \ `__|  / _ \ \/ / __/ _ \ |         
   | | | |  __/\ V  V /  __/ |    |  __/>  < (_|  __/ |         
   |_| |_|\___| \_/\_/ \___|_|     \___/_/\_\___\___|_|         
                                                                
    options nolabel;                                            
                                                                
    %utlfkil(d:\xls\sdtm.xlsx);  * delete;                      
                                                                
    ods excel file="d:\xls\sdtm.xlsx" options(sheet_name="TOC");
                                                                
    proc report data=sashelp.class(obs=1)                       
      style(column)={tagattr='wraptext:no' width=100%};         
          define name   / display  style={color=blue};          
          compute name;                                         
            if name = 'Alfred' then                             
                 call define('name','url',"sdtm.xlsx#DM!A1");   
          endcomp;                                              
    run;quit;                                                   
                                                                
    ods excel options(sheet_name="DM");                         
                                                                
    proc report data=sashelp.class missing;                     
    ;run;quit;                                                  
                                                                
    ods excel close;                                            
    options label;                                              


