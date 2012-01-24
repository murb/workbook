require 'test/helper'

class TestFormat < Test::Unit::TestCase

  
  def test_initialize
    f = Workbook::Format.new {}
    assert_equal({},f)
    f = Workbook::Format.new({:background=>:red})
    assert_equal({:background=>:red},f)
    f = Workbook::Format.new({:background=>:red})
    deet = Date
    f.raw = deet
    f = Workbook::Format.new(f)
    assert_equal(deet,f.raw)
    
    
  end
 
  
end
