# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

class TestTable< Test::Unit::TestCase
  def test_initialize
    t = Workbook::Table.new
    assert_equal(t,[])
    c = Workbook::Cell.new("celllll")
    t = Workbook::Table.new [[c]]
     
    assert_equal([[c]],t)
  end
  def test_header
    t = Workbook::Table.new
    assert_equal(t.header,nil)
    t = Workbook::Table.new [[1]]
    assert_equal(t.header,[1])
    assert_equal(t.header.class,Workbook::Row)
  end
  
  def test_new_row
    t = Workbook::Table.new
    assert_equal(t.count, 0)
    
    r = t.new_row [1,2,3,4]
    assert_equal(r, [1,2,3,4])
    assert_equal(r.class, Workbook::Row)    
    assert_equal(1,t.count)
    
    r = t.new_row
    assert_equal(r.empty?, true)
    assert_equal(r.class, Workbook::Row)
    
    assert_equal(2, t.count)
    assert_equal(r, t.last)
    
    r << 2

    assert_equal(t.last.empty?, false)
  end
  
  def test_append_row
    t = Workbook::Table.new
    row = t.new_row(["a","b"])
    assert_equal(row, t.header)
    row = Workbook::Row.new([1,2])
    assert_equal(nil, row.table)
    t.push(row)
    assert_equal(t, row.table)
    row = Workbook::Row.new([3,4])
    assert_equal(nil, row.table)
    t << row 
    assert_equal(t, row.table)
  end
  
  def test_sheet
    t = Workbook::Table.new
    s = t.sheet
    assert_equal(t, s.table)
    assert_equal(t.sheet, s)
  end
  
  def test_name
    t = Workbook::Table.new
    t.name = "test naam"
    assert_equal("test naam", t.name)
  end
  
  def test_delete_all
    w = Workbook::Book.new [["a","b"],[1,2],[3,4]]
    t = w.sheet.table
    t.delete_all
    assert_equal(Workbook::Table,t.class)
    assert_equal(0,t.count)
  end

  def test_clone
    w = Workbook::Book.new [["a","b"],[1,2],[3,4]]
    t = w.sheet.table
    assert_equal(3,t[2][:a])
    t2 = t.clone
    t2[2][:a] = 5
    assert_equal(3,t[2][:a])
    assert_equal(5,t2[2][:a])
  end
  
  def test_clone_custom_header
    w = Workbook::Book.new [[nil, nil],["a","b"],[1,2],[3,4]]
    t = w.sheet.table
    t.header=t[1]
    assert_equal(3,t[3][:a])
    t2 = t.clone
    t2[3][:a] = 5
    assert_equal(3,t[3][:a])
    assert_equal(5,t2[3][:a])
  end

end
