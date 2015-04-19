# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

class TestColumn < Minitest::Test

  def new_table
    Workbook::Table.new([
        ["a","b","c","d"],
        [true,3.2,"asdf",1],
        [true,3.2,"asdf",1],
        [false,3.2,"asdf",1],
        [true,3.2,"asdf",1]
      ])
  end

  def test_init
    c = Workbook::Column.new
    assert_equal(Workbook::Column, c.class)
    c = Workbook::Column.new(nil, {:limit=>20,:default=>"asdf", :column_type=>:boolean})
    c = Workbook::Column.new(Workbook::Table.new, {:limit=>20,:default=>"asdf", :column_type=>:boolean})
    assert_equal(20, c.limit)
    assert_equal(Workbook::Cell.new("asdf"), c.default)
    assert_equal(:boolean, c.column_type)
    assert_raises(ArgumentError) { Workbook::Column.new(true) }
    assert_raises(ArgumentError) { Workbook::Column.new(nil, {:limit=>20,:default=>"asdf", :column_type=>:bodfolean}) }
  end

  def test_table
    c = Workbook::Column.new
    c.table = Workbook::Table.new
    assert_equal(Workbook::Table.new, c.table)
    assert_raises(ArgumentError) { c.table = false }
  end

  def test_index
    t = new_table
    assert_equal(t.columns.first.index, 0)
    assert_equal(t.columns.last.index, 3)
  end

  def test_column_type
    t = new_table
    assert_equal([:boolean, :float, :string, :integer], t.columns.collect{|a| a.column_type})
    t = new_table
    t.last.last.value = 1.1
    assert_equal([:boolean, :float, :string, :string], t.columns.collect{|a| a.column_type})
  end
end