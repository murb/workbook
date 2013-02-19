# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

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
  
  def test_merge
    a = Workbook::Format.new({:background=>:red})
    b = Workbook::Format.new({:background=>:yellow, :color=>:green})
    result = b.clone.merge(a)
    assert_equal({:background=>:red, :color=>:green},result)
    assert_equal(true,result.is_a?(Workbook::Format))
  end
  
  def test_remove_raw_on_merge
    a = Workbook::Format.new({:background=>:red})
    b = Workbook::Format.new({:background=>:yellow, :color=>:green})
    b.add_raw Date.new
    result = b.clone.merge(a)
    assert_equal({},result.raws)
  end
 
  def test_has_background_color?
    a = Workbook::Format.new
    assert_equal(false,a.has_background_color?)
    a = Workbook::Format.new({:background_color=>'#ff0000'})
    assert_equal(true,a.has_background_color?)
    assert_equal(false,a.has_background_color?('#00ff00'))
    a = Workbook::Format.new({:background_color=>'#ffffff'})
    assert_equal(false,a.has_background_color?)
    a = Workbook::Format.new({:background_color=>'#FFFFFf'})
    assert_equal(false,a.has_background_color?)

    
  end
  
  def test_to_css
    a = Workbook::Format.new({:background_color=>'#ffffff'})
    assert_equal("background: #ffffff",a.to_css)
    a = Workbook::Format.new({:background_color=>'#fffdff', :color=>:red})
    assert_equal("background: #fffdff; color: red",a.to_css)
    
  end
  
end
