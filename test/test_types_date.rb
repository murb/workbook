# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

class TestTypesDate < Minitest::Test
  def test_init
    w = Workbook::Types::Date.new(2001,2,2)
    assert_equal(Workbook::Types::Date, w.class)
    assert_equal(Date.new(2001,2,2),w)
    assert_equal(Date.new(2001,2,2),w.value)
  end
end