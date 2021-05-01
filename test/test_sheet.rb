# frozen_string_literal: true

require File.join(File.dirname(__FILE__), "helper")

class TestWorkbook < Minitest::Test
  def test_init
    s = Workbook::Sheet.new nil
    assert_equal(Workbook::Sheet, s.class)
    assert_equal(s.count, 1)

    s = Workbook::Sheet.new
    assert_equal(Workbook::Sheet, s.class)
    assert_equal(s.count, 1)

    t = Workbook::Table.new []
    s = Workbook::Sheet.new t
    assert_equal(Workbook::Sheet, s.class)
    assert_equal(t, s.table)
    assert_equal(s.count, 1)
  end

  def test_table
    w = Workbook::Sheet.new nil
    assert_equal([], w.table)
    t = Workbook::Table.new []
    w = Workbook::Sheet.new t
    assert_equal(w.table, t)
  end

  def test_table_assignment
    t = Workbook::Table.new []
    s = Workbook::Sheet.new t
    assert_equal(s.table, t)

    data = [["a", "b"], [1, 2]]
    s.table = data
    assert_equal("a", s.table["A1"].value)
    assert_equal(2, s.table["B2"].value)
  end

  def test_book
    s = Workbook::Sheet.new
    b = s.book
    assert_equal(s.book, b)
    assert_equal(s, b.sheet)
    assert_equal(s.book.sheet, b.sheet.table.sheet)
  end

  def test_clone
    w = Workbook::Book.new [["a", "b"], [1, 2], [3, 4]]
    s = w.sheet

    assert_equal(3, s.table[2][:a])

    s2 = s.clone

    s2.table[2][:a] = 5
    assert_equal(3, s.table[2][:a])
    assert_equal(5, s2.table[2][:a])
  end

  def test_create_or_open_table_at
    s = Workbook::Sheet.new
    table0 = s.create_or_open_table_at(0)
    assert_equal(Workbook::Table, table0.class)
    assert_equal(s, table0.sheet)
    table1 = s.create_or_open_table_at(1)
    assert_equal(Workbook::Table, table1.class)
    assert_equal(s, table1.sheet)
    table1 << Workbook::Row.new([1, 2, 3, 4])
    assert_equal(false, table1 == table0)
  end

  def test_profile_speed
    w = Workbook::Book.new [["a", "b"], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4], [1, 2], [3, 4]]
    require "ruby-prof"
    RubyProf.start
    w.sheet.table.each do |row|
      row[:a].value
    end
    result = RubyProf.stop
    printer = RubyProf::MultiPrinter.new(result)
    printer.print(path: ".", profile: "profile")
  end

  def test_name
    b = Workbook::Book.new [["a", "b"], [1, 2]]
    b.push Workbook::Sheet.new([["a", "b"], [2, 2]])
    b.push Workbook::Sheet.new([["a", "b"], [3, 2]])

    # puts b.index b.last
    assert_equal(["Sheet 1", "Sheet 2", "Sheet 3"], b.collect { |a| a.name })
  end
end
