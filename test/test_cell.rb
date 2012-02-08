require 'test/helper'

class TestCell < Test::Unit::TestCase

  
  def test_init
    w = Workbook::Cell.new nil
    assert_equal(nil,w.value)
    w = Workbook::Cell.new "asdf"
    assert_equal("asdf",w.value)
    
    assert_raise(ArgumentError) { w = Workbook::Cell.new :asdf }

    t = Time.now
    w = Workbook::Cell.new t
    assert_equal(t,w.value)
    
  end
 
  def test_value
    w = Workbook::Cell.new nil
    assert_equal(nil,w.value)
    w.value = "asdf"
    assert_equal("asdf",w.value)
    w.value = Date.new
    assert_equal(Date.new,w.value) 
  end
  
end
