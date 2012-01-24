require 'test/helper'

class TestWorkbook < Test::Unit::TestCase
  def test_sheets
    w = Workbook::Book.new nil
    assert_equal(w,[])
    w.push
    assert_equal(w.count,1)
  end
  
  def test_push
    w = Workbook::Book.new nil
    assert_equal(w,[])
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
  
end
