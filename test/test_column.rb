# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

class TestRow < Test::Unit::TestCase

  def test_init
    c = Workbook::Column.new
    assert_equal(Workbook::Column, c.class)
    c = Workbook::Column.new({:limit=>20,:default=>"asdf", :column_type=>:boolean})
    assert_equal(20, c.limit)
    assert_equal(Workbook::Cell.new("asdf"), c.default)
    assert_equal(:boolean, c.column_type)
    assert_raise(ArgumentError) { Workbook::Column.new({:limit=>20,:default=>"asdf", :column_type=>:bodfolean}) }
  end
end