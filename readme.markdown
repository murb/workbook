# Workbook

Workbook is a gem that mimicks a typical spreadsheet, a bundle of sheets, bundled in a *workbook*. A sheet may contain one or more tables (which might  the multi table sheets of Apple Numbers or Excel ranges). Book, Sheet, Table and Row inherit from the base Array class, and hence walks and quacks as such. Values are converted to ruby native types. 

Goals of this gem:

* [Done] No semantic DSL approach (if want to try a DSL-approach for creating sheets try [https://github.com/kellyredding/osheet/wiki](OSheet)), but instead try to stay as close to the original Ruby language as possible. “A 2-D array is a good basis for a table-store”. You should not have to learn how this plugin works.
* [Done] Allow for standard Array and Hash operations
* [Done] Make it easy to diff two tables
* [Done] Make it easy to import tables, parsing values to matching ruby-types and proper unicode characters (currently XLS (`spreadsheet-gem`), CSV (`faster_csv`) and TXT (`faster_csv`) support provided)
* Write excel files based on template-files
* Usability over speed.

## Hierarchy of concepts

* Book
   * Sheet 
      * Table
        
Subsequently a table consists of:

* Table
   * Row
      * Cell ( wich has may have a (shared) Format )
	  
## Initializing
	  
Simply initialize a simple spreadsheet using:

    b = Workbook::Book.new
	
or

    b = Workbook::Book.open filename
	   
Calling

    s = b.sheet
	t = s.table
	
will give you an empty, or the first, Sheet and Table.

You may want to initialize the whole shebang from a 2-d array, like this:

    b = Workbook::Book.new [['a','b'],[1,2],[3,4],[5,6]]
	t = s.sheet.table
	
Subsequently you lookup values in the table like this:

    t[1][:b] 
	# returns <Workbook::Cel @value=2>
	
Returning a Workbook::Cel is a bit of a quick fix, ideally I would just serve you the right value. While ruby has the powers to offer that properly, retaining a reference to the original raw cell content and conveying a bit of preferred formatting, I simply haven't undertaken the effort to implement this properly.
	
<!-- Feature *to implement*: 

	t['A2']
	# returns <Workbook::Cel @value=1>
	
Feature *to implement*, get a single column:

    t[:b]
	# returns [<Workbook::Cel @value=2>,<Workbook::Cel @value=4>,<Workbook::Cel @value=6>] -->
	
## Sorting & Comparing

Sorting leaves the header alone, and doesn't complain about comparing strings with dates with floats. We're talking spreadsheet here. When classes differ the following (default) order is used: Numbers, Strings, Dates and Times, Booleans and Nils (empty values).

	t.sort
	
*To implement*:

To some extent, sort_by works, it doesn't, however, adhere to the header settings... 
  
    t.sort_by {|r| r[:b]}
	
But sorting was implemented to enable easy comparison between tables. Simply call 

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