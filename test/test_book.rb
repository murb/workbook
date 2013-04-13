# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

class TestWorkbook < Test::Unit::TestCase
  def test_sheets
    w = Workbook::Book.new nil
    w.push
    assert_equal(2, w.count)
  end
  
  def test_push
    w = Workbook::Book.new nil
    assert_equal([[[]]],w)
    w = Workbook::Book.new
    assert_equal(w.count,1)
    
    w.push
    assert_equal(w.first.class,Workbook::Sheet)
    w.push
    assert_equal(w.count,3)    
    s = Workbook::Sheet.new
    w.push s
    assert_equal(w.last,s)
    w = Workbook::Book.new
    assert_equal(w.sheet.table.class,Workbook::Table)
  end
  
  def test_sheet
    w = Workbook::Book.new nil
    s = Workbook::Sheet.new [Workbook::Row.new(Workbook::Table.new)]
    assert_equal(w.sheet.class,Workbook::Sheet)
    assert_not_equal(w.sheet, s)
    w = Workbook::Book.new s
    assert_equal(w.sheet, s)
  end
  
  def test_template
    b = Workbook::Book.new
    raw = "asdf"
    assert_raise(ArgumentError) { b.template = raw }
    raw = Workbook::Template.new
    b.template = raw
    
    assert_equal(raw,b.template)
  end
  
  def test_parent_child
    b = Workbook::Book.new [[1,2,3],[1,2,3]]
    assert_equal(Workbook::Sheet, b.first.class)
    assert_equal(b,b.first.book)
    assert_equal(Workbook::Table, b.first.table.class)
    assert_equal(b,b.first.table.sheet.book)
    assert_equal(Workbook::Row, b.first.table.header.class)
    assert_equal(b,b.first.table.header.table.sheet.book)
  end
  
  def test_text_to_utf8
    f = File.open("test/artifacts/excel_different_types.txt",'r')
    t = f.read
    w = Workbook::Book.new
    t = w.text_to_utf8(t)
    assert_equal("a\tb\tc\td", t.split(/(\n|\r)/).first)
  end
  
  def test_read_bad_filetype
    assert_raises(ArgumentError) { Workbook::Book.read("test string here", :xls) }
    assert_raises(ArgumentError) { Workbook::Book.read("test string here", :ods) }
    assert_raises(ArgumentError) { Workbook::Book.read("test string here", :xlsx) }
  end
end
