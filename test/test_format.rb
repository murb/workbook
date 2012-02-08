require 'test/helper'

class TestFormat < Test::Unit::TestCase

  
  def test_initialize
    f = Workbook::Format.new {}
    assert_equal({},f)
    f = Workbook::Format.new({:background=>:red})
    assert_equal({:background=>:red},f)
    f = Workbook::Format.new({:background=>:red})
    deet = Time.now
    assert_equal(false,f.has_raw_for?(Time))
    
    f.add_raw deet
    assert_equal(deet,f.raws[Time])
    assert_equal(true,f.has_raw_for?(Time))
    
    
  end
 
  
end
