# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

class TestColumn < Test::Unit::TestCase

  def test_init
    c = Workbook::Column.new
    assert_equal(Workbook::Column, c.class)
    c = Workbook::Column.new(nil, {:limit=>20,:default=>"asdf", :column_type=>:boolean})
    c = Workbook::Column.new(Workbook::Table.new, {:limit=>20,:default=>"asdf", :column_type=>:boolean})
    assert_equal(20, c.limit)
    assert_equal(Workbook::Cell.new("asdf"), c.default)
    assert_equal(:boolean, c.column_type)
    assert_raise(ArgumentError) { Workbook::Column.new(true) }
    assert_raise(ArgumentError) { Workbook::Column.new(nil, {:limit=>20,:default=>"asdf", :column_type=>:bodfolean}) }
  end

  def test_table
    c = Workbook::Column.new
    c.table = Workbook::Table.new
    assert_equal(Workbook::Table.new, c.table)
    assert_raise(ArgumentError) { c.table = false }
  end
end