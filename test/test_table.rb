require 'test/helper'

class TestTable< Test::Unit::TestCase
  def test_initialize
    t = Workbook::Table.new
    assert_equal(t,[])
    t = Workbook::Table.new [[1]]
    assert_equal(t,[[1]])
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
    assert_equal(t.count, 1)
    
    r = t.new_row
    assert_equal(r.empty?, true)
    assert_equal(r.class, Workbook::Row)
    
    assert_equal(t.count, 2)
    assert_equal(t.last, r)
    
    r << 2
    
    assert_equal(t.last.empty?, false)
    
    
  end

end
