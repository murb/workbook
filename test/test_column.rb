# frozen_string_literal: true

require File.join(File.dirname(__FILE__), "helper")

class TestColumn < Minitest::Test
  def new_table
    Workbook::Table.new([
      ["a", "b", "c", "d"],
      [true, 3.2, "asdf", 1],
      [true, 3.2, "asdf", 1],
      [false, 3.2, "asdf", 1],
      [true, 3.1, "asdf", 1]
    ])
  end

  def test_init
    c = Workbook::Column.new
    assert_equal(Workbook::Column, c.class)
    c = Workbook::Column.new(Workbook::Table.new, {limit: 20, default: "asdf", column_type: :boolean})
    assert_equal(20, c.limit)
    assert_equal(Workbook::Cell.new("asdf"), c.default)
    assert_equal(:boolean, c.column_type)
    assert_raises(ArgumentError) { Workbook::Column.new(true) }
    assert_raises(ArgumentError) { Workbook::Column.new(nil, {limit: 20, default: "asdf", column_type: :bodfolean}) }
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
    assert_equal([:boolean, :float, :string, :integer], t.columns.collect { |a| a.column_type })
    t = new_table
    t.last.last.value = 1.1
    assert_equal([:boolean, :float, :string, :string], t.columns.collect { |a| a.column_type })
    t = new_table
    t[2][3] = nil
    assert_equal([:boolean, :float, :string, :integer], t.columns.collect { |a| a.column_type })
    t = new_table
    t[2].delete_at(3)
    assert_equal([:boolean, :float, :string, :integer], t.columns.collect { |a| a.column_type })
  end

  def test_alpha_index_to_number_index
    assert_equal(0, Workbook::Column.alpha_index_to_number_index("A"))
    assert_equal(2, Workbook::Column.alpha_index_to_number_index("C"))
    assert_equal(25, Workbook::Column.alpha_index_to_number_index("Z"))
    assert_equal(26, Workbook::Column.alpha_index_to_number_index("AA"))
    assert_equal(27, Workbook::Column.alpha_index_to_number_index("AB"))
    assert_equal(51, Workbook::Column.alpha_index_to_number_index("AZ"))
    assert_equal(52, Workbook::Column.alpha_index_to_number_index("BA"))
    assert_equal((27 * 26) - 1, Workbook::Column.alpha_index_to_number_index("ZZ"))
  end

  def test_cells
    assert_equal([3.2,3.2,3.2,3.1], new_table.columns[1].cells.map(&:value))
    assert_equal([3.2,3.2,3.2,3.1], new_table.columns[1].map(&:value))
  end
end
