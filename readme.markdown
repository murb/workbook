# Workbook

Goal of this gem is to make working with workbooks (spreadsheets) as programmer friendly as possible. Not reinventing a totally new DSL or all kinds of new methodnames, but just borrowing from known concepts such as hashes and arrays (much like (Faster)CSV does)). Workbook is a gem that mimicks a typical spreadsheet, a bundle of sheets, bundled in a *workbook*. A sheet may contain one or more tables (which might  the multi table sheets of Apple Numbers or Excel ranges). Basically:

* Book
   * Sheet (one or more)
      * Table (one or more)
        
Subsequently a table consists of:

* Table
   * Row (one or more)
      * Cell ( wich has may have a (shared) Format )
	  
Book, Sheet, Table and Row inherit from the base Array class, and hence walks and quacks as such. The row is extended with hashlike lookups (`row[:id]`) and writers (`row[:id]=`). Values are converted to ruby native types, and optional parsers can be added to improve recognition. 

In addition to offering you this plain structure it allows for importing and writing .xls and .csv files (more to come), and includes the utility to easily create an overview of the differences between two tables and read out basic cell-styling properties as css.

## The Basics
	  
Simply initialize a simple spreadsheet using:

    b = Workbook::Book.new
	
or

    b = Workbook::Book.open filename
	   
Calling

    s = b.sheet
	t = s.table
	
will give you an the first Sheet and Table (if one doesn't exist it is created on the fly).

You can initialize with simple 2-d array like this:

    b = Workbook::Book.new [['a','b'],[1,2],[3,4],[5,6]]
	t = s.sheet.table
	
Subsequently you lookup values in the table like this:

    t[1][:b] 
	# returns <Workbook::Cel @value=2>
	

<!-- Feature *to implement*: 

	t['A2']
	# returns <Workbook::Cel @value=1>
	
Feature *to implement*, get a single column:

    t[:b]
	# returns [<Workbook::Cel @value=2>,<Workbook::Cel @value=4>,<Workbook::Cel @value=6>] 
	
On my wishlist: In the future I hope to return the cell value directly, without the intermediate Workbook::Cel class in between.
	
	-->
	
## Utilities

### Sorting

Sorting leaves the header alone, if it exists, and doesn't complain about comparing strings with dates with floats (Ever found OpenOffice Calc or Excel complainging about its inability to compare integers and strings? We're talking spreadsheet here). When classes are different the following (default) order is used: Numbers, Strings, Dates and Times, Booleans and Nils (empty values).

	t.sort
	
*To implement*:

To some extent, sort_by works, it doesn't, however, adhere to the header settings... 
  
    t.sort_by {|r| r[:b]}
	
### Comparing tables
	
Simply call 

	t1.diff t2
	
And a new book with a new sheet/table will be returned containing the differences between the two tables.
	
## Writing

Currently writing is limited to the following formats. Templating support is still limited.
	
	b.to_xls 					# returns a spreadsheet workbook
	b.write_to_xls filename 	# writes to filename
	t.to_csv 					# returns a csv-string
	
In case you want to display the table in HTML, some conversion is offered to convert text/background properties to css-entities. Internally the hash storing style elements tries to map to CSS where possible.
	
## Alternatives

The [ruby toolbox lists plenty of alternatives](https://www.ruby-toolbox.com/search?utf8=%E2%9C%93&q=spreadsheet), that just didn't suit my needs.

## License

MIT... (c) murb / Maarten Brouwers, 2011-2012