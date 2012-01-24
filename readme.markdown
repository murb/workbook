# Workbook

Workbook is a gem that mimicks a typical spreadsheet, a bundle of sheets, bundled in a *workbook*. A sheet may contain one or more tables (which might  the multi table sheets of Apple Numbers or Excel ranges) 

* book
   * sheet 
      * table
         * row
            * cell

Simply initialize a simple spreadsheet using:

    b = Workbook::Book.new
   
Calling

    s = b.sheet
	
will give you an empty Sheet.
