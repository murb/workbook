# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

class TestWorkbook < Test::Unit::TestCase
  def test_init
    w = Workbook::Sheet.new nil
    assert_equal([[]],w)
    assert_equal(w.count,1)
    w = Workbook::Sheet.new
    assert_equal([Workbook::Table.new],w)
    assert_equal(w.count,1)
    t = Workbook::Table.new []
    w = Workbook::Sheet.new t
    assert_equal([t],w)    
    assert_equal(w.count,1)
  end
  
  def test_table
    w = Workbook::Sheet.new nil
    assert_equal([],w.table)
    t = Workbook::Table.new []
    w = Workbook::Sheet.new t
    assert_equal(w.table,t)
    
  end
  
  def test_book
    s = Workbook::Sheet.new
    b = s.book
    assert_equal(s.book, b)
    assert_equal(s, b.sheet)
    assert_equal(s.book.sheet, b.sheet.table.sheet)
  end
  
  def test_clone
    w = Workbook::Book.new [["a","b"],[1,2],[3,4]]
    s = w.sheet
    assert_equal(3,s.table[2][:a])
    s2 = s.clone
    s2.table[2][:a] = 5
    assert_equal(3,s.table[2][:a])
    assert_equal(5,s2.table[2][:a])
  end
 
end
