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

end
