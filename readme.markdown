# Workbook

Workbook is a gem that mimicks a typical spreadsheet, a bundle of sheets, bundled in a *workbook*. A sheet may contain one or more tables (which might  the multi table sheets of Apple Numbers or Excel ranges). Book, Sheet, Table and Row inherit from the base Array class, and hence walks and quacks as such. 

Goals of this gem:

* [Done] No semantic DSL approach (if want to try a DSL-approach for creating sheets try [https://github.com/kellyredding/osheet/wiki](OSheet)), but instead try to stay as close to the original Ruby language as possible. “A 2-D array is a good basis for a table-store”. 
* [Done] Allow for standard Array and Hash operations
* Make it easy to sort values in columns
* Make it easy to diff two tables
* Make it possible to maintain links to 'proprietary' objects, such as produced by the spreadsheet gem, hence allowing for using existing files as 'templates'.

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
   
Calling

    s = b.sheet
	t = s.table
	
will give you an empty Sheet and Table.

You may want to initialize the whole shebang from a 2-d array, like this:

    b = Workbook::Book.new
    s = b.sheet[0] = Workbook::Sheet.new([['a','b'],[1,2],[3,4],[5,6]])
	t = s.table
	
Subsequently you lookup values in the table like this:

    t[1][:b] 
	# returns <Workbook::Cel @value=2>
	
Feature *to implement*: 

	t['A2']
	# returns <Workbook::Cel @value=1>
	
Feature *to implement*:

    t[:b]
	# returns [<Workbook::Cel @value=2>,<Workbook::Cel @value=4>,<Workbook::Cel @value=6>]
	
## Sorting

Sorting leaves the header alone, and doesn't (shouldn't) complain about comparing strings with dates with floats. We're talking spreadsheet here.

*To implement*:
  
    t.sort_by {|r| r[:b]}
	
## Writing

*To implement*:
	
	t.to_xls
	s.to_xls
	b.to_xls
	t.to_csv
	b.to_xlsx
	
## Alternatives

The [ruby toolbox lists plenty of alternatives](https://www.ruby-toolbox.com/search?utf8=%E2%9C%93&q=spreadsheet), that just didn't suit my needs.

## License

MIT... (c) murb / Maarten Brouwers, 2011