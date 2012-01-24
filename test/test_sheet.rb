require 'test/helper'

class TestWorkbook < Test::Unit::TestCase
  def test_init
    w = Workbook::Sheet.new nil
    assert_equal([nil],w)
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
    assert_equal(w.table,nil)
    t = Workbook::Table.new []
    w = Workbook::Sheet.new t
    assert_equal(w.table,t)
    
  end
 
end
