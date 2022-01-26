# frozen_string_literal: true

require File.join(File.dirname(__FILE__), "helper")

class TestTable < Minitest::Test
  def test_initialize
    t = Workbook::Table.new
    assert_equal([], t.rows)
    c = Workbook::Cell.new("celllll")
    t = Workbook::Table.new [[c]]

    assert_equal([[c]], t.rows.map(&:cells))
  end

  def test_header
    t = Workbook::Table.new
    assert_nil(t.header)
    t = Workbook::Table.new [[1]]
    assert_equal(t.header, [1])
    assert_equal(t.header.class, Workbook::Row)
  end

  def test_new_row
    t = Workbook::Table.new
    assert_equal(t.count, 0)

    r = t.new_row [1, 2, 3, 4]
    assert_equal(r, [1, 2, 3, 4])
    assert_equal(r.class, Workbook::Row)
    assert_equal(1, t.count)

    r = t.new_row
    assert_equal(r.empty?, true)
    assert_equal(r.class, Workbook::Row)

    assert_equal(2, t.count)

    assert_equal(r, t.rows.last)
    assert_equal(r, t.last)

    r << 2

    assert_equal(t.last.empty?, false)
  end

  def test_append_row
    t = Workbook::Table.new
    row = t.new_row(["a", "b"])
    assert_equal(row, t.header)
    row = Workbook::Row.new([1, 2])
    assert_nil(row.table)
    t.push(row)
    assert_equal(t, row.table)
    row = Workbook::Row.new([3, 4])
    assert_nil(row.table)
    t << row
    assert_equal(t, row.table)
    t = Workbook::Table.new
    t << [1, 2, 3, 4]
    assert_equal(Workbook::Row, t.first.class)
  end

  def test_sheet
    t = Workbook::Table.new
    s = t.sheet
    assert_equal(t, s.table)
    assert_equal(t.sheet, s)
  end

  def test_name
    t = Workbook::Table.new
    t.name = "test naam"
    assert_equal("test naam", t.name)
  end

  def test_delete_all
    w = Workbook::Book.new [["a", "b"], [1, 2], [3, 4]]
    t = w.sheet.table
    t.delete_all
    assert_equal(Workbook::Table, t.class)
    assert_equal(0, t.count)
  end

  def test_clone
    w = Workbook::Book.new [["a", "b"], [1, 2], [3, 4]]
    t = w.sheet.table
    assert_equal(3, t[2][:a])
    t2 = t.clone
    t2[2][:a] = 5
    assert_equal(3, t[2][:a])
    assert_equal(5, t2[2][:a])
  end

  def test_clone_custom_header
    w = Workbook::Book.new [[nil, nil], ["a", "b"], [1, 2], [3, 4]]
    t = w.sheet.table
    t.header = t[1]
    assert_equal(3, t[3][:a])
    t2 = t.clone
    assert_equal(1, t2.header_row_index)

    t2[3][:a] = 5
    assert_equal(3, t[3][:a])
    assert_equal(5, t2[3][:a])
  end

  def test_spreadsheet_style_cell_addressing
    w = Workbook::Book.new [[nil, nil], ["a", "b"], [1, 2], [3, 4]]
    t = w.sheet.table
    assert_nil(t["A1"].value)
    assert_nil(t["B1"].value)
    assert_equal("a", t["A2"].value)
    assert_equal("b", t["B2"].value)
    assert_equal(1, t["A3"].value)
    assert_equal(2, t["B3"].value)
    assert_equal(3, t["A4"].value)
    assert_equal(4, t["B4"].value)
  end

  def test_multirowselect_through_collections
    w = Workbook::Book.new [["a", "b"], [1, 2], [3, 4]]
    t = w.sheet.table
    assert_equal(Workbook::Table, t[0..2].class)
    assert_equal(2, t[0..2][1][1])
  end

  def test_table
    w = Workbook::Book.new [[nil, nil], ["a", "b"], [1, 2], [3, 4]]
    t = w.sheet.table
    w2 = Workbook::Book.new
    w2.sheet.table = t[2..3]
    assert_equal("1,2\n3,4\n", w2.sheet.table.to_csv)
  end

  def test_array_style_assignment
    w = Workbook::Book.new [["a", "b"], [1, 2], [3, 4]]
    t = w.sheet.table
    r = t[1].clone
    assert_nil(r.table)
    t[2] = r
    assert_equal(t, r.table)
  end

  def test_delete_at
    w = Workbook::Book.new [["a", "b"], [1, 2], [3, 4]]
    t = w.sheet.table
    t.delete_at 2
    assert_equal(1, t.last.first.value)
  end

  def test_trim!
    t = Workbook::Table.new
    t << [1, 2, 3]
    t << [1, 2, nil, nil]
    t.trim!
    assert_equal("1,2,3\n1,2,\n", t.to_csv)
    t = Workbook::Table.new
    t << [1, 2, 3]
    t << [nil]
    t << [1, 2, nil, nil]
    t << [nil, nil, nil, nil]
    t << [nil, nil, nil, nil]
    t.trim!
    assert_equal("1,2,3\n,,\n1,2,\n", t.to_csv)
  end

  def test_performance
    table = Workbook::Table.new
    headers = 100.times.collect { |a| "header#{a}" }
    first_row = 100.times.collect { |a| Time.now }
    table << headers.shuffle
    table << first_row
    1000.times do |times|
      row = table[1].clone
      table << row
      headers.each do |a|
        row[a.to_sym] = Time.now
      end
    end
    last_line = table.count - 1
    delta_start = table[12][0].value - table[2][0].value
    delta_end = table[last_line][0].value - table[last_line - 10][0].value
    average_run_time = (delta_start + delta_end / 20)
    if (delta_end - delta_start) > average_run_time
      puts "Performance issue"
    end
  end

  def test_columns
    table = Workbook::Table.new([[]])
    assert_equal(table.columns, [])
    table = Workbook::Table.new([[:a, :b], [1, 2]])
    assert_equal(table.columns.count, 2)
  end

  def test_dimensions
    table = Workbook::Table.new([[]])
    assert_equal([0, 1], table.dimensions)
    table = Workbook::Table.new([[:a, :b], [1, 2, 3, 4]])
    assert_equal([4, 2], table.dimensions)
  end

  def test_push
    table = Workbook::Table.new([["a", 2, 3]])
    table.push(["b", 2, 3])
    assert_equal("b", table[1][0].value)
    table.push(["c", 2, 3], ["d", 2, 3])
    assert_equal("d", table[3][0].value)
  end
end
