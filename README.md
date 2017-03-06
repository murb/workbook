# Workbook
[![Code Climate](https://codeclimate.com/github/murb/workbook.png)](https://codeclimate.com/github/murb/workbook) [![Build Status](https://travis-ci.org/murb/workbook.svg?branch=master)](https://travis-ci.org/murb/workbook) [![Gem Version](https://badge.fury.io/rb/workbook.svg)](http://badge.fury.io/rb/workbook)

Goal of this gem is to make working with workbooks (spreadsheets) as programmer friendly as possible. Not reinventing a totally new DSL or all kinds of new methodnames, but just borrowing from known concepts such as hashes and arrays (much like (Faster)CSV does)). Workbook is a gem that mimicks a typical spreadsheet, a bundle of sheets, bundled in a *workbook*. A sheet may contain one or more tables (which might the multi table sheets of Apple Numbers or Excel ranges). Basically:

* Book
   * Sheet (one or more)
      * Table (one or more)

Subsequently a table consists of:

* Table
   * Row (one or more)
      * Cell ( wich has may have a (shared) Format )

Book, Sheet, Table and Row inherit from the base Array class, and hence walks and quacks as such. The row is extended with hashlike lookups (`row[:id]`) and writers (`row[:id]=`). Values are converted to ruby native types, and optional parsers can be added to improve recognition.

In addition to offering you this plain structure it allows for importing .xls, .csv, .xlsx, .txt files (more to come), writing .xls, and .csv  (more to come) and includes several utilities to easily create an overview of the differences between two tables and output basic cell-styling properties as css.

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

Subsequently you look up values in the table like this:

    t[1][:b]
    # returns <Workbook::Cel @value=2>

which is equivalent to

    t[1][1]

Of course you'll be able to write a new value back to it. If you just enter a value, formatting of the original cell will be maintained.

    t[1][:b] = 5

Alternatively (more spreadsheet like) you can read cells like this (writing to be supported, not implemented yet)

    t['A2']

If you want to use an existing file as a template (which you can create in Excel to create nice looking templates),
simply clone the row, and add it back:

    b = Workbook::Book.open("template.xls")
    table = b.sheet.table
    template_row = table[1]            # can be any, but I typically have a well
                                    # formatted header row + an example template
                                    # row for the data
    [1,2,3,4].each do |v|
      new_row = template_row.clone
      table << new_row              # to use the symbol style header references,
                                    # the row first needs to be added back to the
                                    # table
      new_row[:a] = v
    end
    table.delete(template_row)      # you don't want the template to show up
                                    # in the endresult
    b.write("result.xls")           # write it!

Another typical use case is exporting a list of ActiveRecord-objects to xls (it is assumed that the headers of the excel-table correspond
(like "Total order price" and `total_order_price` match) to the headers of the database-table ):

    b = Workbook::Book.open("template.xls")
    table = b.sheet.table
    template_row = table[1]         # see above
    Order.where("created_at > ?", Time.now - 1.week).each do |order|
      new_row = template_row.clone
      new_row.table = table
      order.to_hash.each{|k,v| row[k]=v}
    end
    table.delete(template_row)      # you don't want the template to show up
                                    # in the endresult
    b.write("recent_orders.xls")    # write it!

## Utilities

### Sorting

Sorting leaves the header alone, if it exists, and doesn't complain about comparing strings with dates with floats (Ever found OpenOffice Calc or Excel complainging about its inability to compare integers and strings? We're talking spreadsheet here). When classes are different the following (default) order is used: Numbers, Strings, Dates and Times, Booleans and Nils (empty values).

    t.sort

*To implement*:

To some extent, sort_by works, it doesn't, however, adhere to the header settings...

    t.sort_by {|r| r[:b]}

### Comparing tables or entire workbooks

Simply call on a Workbook::Table

	t1.diff t2

And a new book with a new table will be returned containing the differences between the two tables.

Alternatively you can run the same command on workbooks, which will compare sheet by sheet and return a new Workbook

## Writing

Currently writing is limited to the following formats. Templating support is still limited.

    b.to_xls                  # returns a spreadsheet workbook
    b.write_to_xls filename   # writes to filename
    t.(write_)to_csv          # returns a csv-string (called on tables)
    b.(write_)to_html         # returns a clean html-page with all tables; unformatted, format-names are used in the classes
    t.(write_)to_json         # returns the values of a table in json
    t.(write_)to_xlsx         # returns/writes using RubyXL to XLS (unstable, work in progress)

In case you want to display a formatted table in HTML, some conversion is offered to convert text/background properties to css-entities. Internally the hash storing style elements tries to map to CSS where possible.

## Compatibility

Workbook is automatically tested for ruby 1.9, 2.0 and 2.1. Most of it works with 1.8.7 and jruby but not all tests give equal results.
Check [Travis for Workbook's current build status](https://travis-ci.org/murb/workbook) [![Build Status](https://travis-ci.org/murb/workbook.svg?branch=master)](https://travis-ci.org/murb/workbook).

## Future

* Column support, e.g. t[:b] could then return Workbook::Column<[<Workbook::Cel @value=2>,<Workbook::Cel @value=4>,<Workbook::Cel @value=6>]>
* In the future I hope to return the cell value as inheriting from the original value's class, so you don't have to call #value as often.
* xlsx support definitly needs to be improved. Especially template based support.

## Alternatives

The [ruby toolbox lists plenty of alternatives](https://www.ruby-toolbox.com/search?utf8=%E2%9C%93&q=spreadsheet), that just didn't suit my needs.

## License

This code MIT (but see below) © murb / Maarten Brouwers, 2011-2015

Workbook uses the following gems:

* [Spreadsheet](https://github.com/zdavatz/spreadsheet) Used for reading and writing old style .xls files (Copyright © 2010 ywesee GmbH (mhatakeyama@ywesee.com, zdavatz@ywesee.com); GPL3 (License required for closed implementations))
* [ruby-ole](http://code.google.com/p/ruby-ole/) Used in the Spreadsheet Gem (Copyright © 2007-2010 Charles Lowe; MIT)
* [FasterCSV](http://fastercsv.rubyforge.org/) Used for reading CSV (comma separated text) and TXT (tab separated text) files (Copyright © James Edward Gray II; GPL2 & Ruby License)
* [rchardet](http://rubyforge.org/projects/rchardet) Used for detecting encoding in CSV and TXT importers (Copyright © JMHodges; LGPL)
* [axslx](https://github.com/randym/axlsx) Used for writing the newer .xlsx files (with formatting) (Copyright © 2011, 2012 Randy Morgan, MIT License)
* [Nokogiri](http://nokogiri.org/) Used for reading ODS documents (Copyright © 2008 - 2012 Aaron Patterson, Mike Dalessio, Charles Nutter, Sergio Arbeo, Patrick Mahoney, Yoko Harada; MIT Licensed)