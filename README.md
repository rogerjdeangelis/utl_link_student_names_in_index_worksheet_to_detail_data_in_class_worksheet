# utl_link_student_names_in_index_worksheet_to_detail_data_in_class_worksheet
Link student names in index worksheet to detail data in class worksheet    Hyperlink to the excel specific cells.    Two Solutions.     1. Libname.     2. ods excel and proc report.     Just output hyperlinks using.     name=cats('=HYPERLINK("[class.xlsx]class!A',put(seq,2.),'","',name,'")');
