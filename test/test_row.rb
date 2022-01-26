# frozen_string_literal: true

require File.join(File.dirname(__FILE__), "helper")

class TestRow < Minitest::Test
  def test_init
    t = Workbook::Table.new
    r = Workbook::Row.new([1, 2, 3], t)
    assert_equal([1, 2, 3], r.collect { |c| c.value })

    # t = Workbook::Table.new
    c1 = Workbook::Cell.new(1)
    c2 = Workbook::Cell.new(2)
    c3 = Workbook::Cell.new(3)

    r = Workbook::Row.new([c1, c2, c3])

    assert_equal([c1, c2, c3], r.cells)
    assert_equal(Workbook::Row, r.class)
  end

  def test_table=
    r = Workbook::Row.new
    assert_raises(ArgumentError, "table should be a Workbook::Table (you passed a String)") { r.table = "asdf" }
    r.table = nil
    assert_nil(r.table)
    r = Workbook::Row.new

    t = Workbook::Table.new
    r.table = t
    assert_equal(r.table, t)
  end

  def test_header?
    t = Workbook::Table.new
    r1 = Workbook::Row.new
    r1.table = t
    assert_equal(true, r1.header?)
    r2 = Workbook::Row.new
    r2.table = t
    assert_equal(false, r2.header?)
    r2 = Workbook::Row.new
    r2.table = t
    assert_equal(false, r2.header?)
    assert_equal(true, t.first.header?)
    t.header = r2
    assert_equal(true, r2.header?)
    assert_equal(false, t.first.header?)

    assert_equal(r1, t.first)
  end

  def test_first?
    t = Workbook::Table.new
    r1 = Workbook::Row.new
    r1.table = t
    assert_equal(true, r1.first?)
    r2 = Workbook::Row.new
    r2.table = t
    assert_equal(false, r2.first?)
    assert_equal(true, t.first.first?)

    assert_equal(r1, t.first)
  end

  def test_no_values?
    t = Workbook::Table.new
    r1 = Workbook::Row.new
    r1.table = t
    assert_equal(true, r1.no_values?)
    r1 << Workbook::Cell.new("abcd")
    assert_equal(false, r1.no_values?)
    r2 = Workbook::Row.new [nil, "", nil, "", ""]
    r2.table = t
    assert_equal(true, r2.no_values?)
  end

  def test_to_symbols
    r1 = Workbook::Row.new ["test", "asdf-asd", "asdf - asdf", "asdf2"]
    assert_equal([:test, :asdfasd, :asdf_asdf, :asdf2], r1.to_symbols)
    r1 = Workbook::Row.new ["inït", "è-éë"]
    assert_equal([:init, :eee], r1.to_symbols)
  end

  def test_to_hash
    r1 = Workbook::Row.new ["test", "asdf-asd", "asdf - asdf", "asdf2"]
    assert_raises(NoMethodError, "undefined method `header' for nil:NilClass") { r1.to_hash }

    t = Workbook::Table.new
    r1 = Workbook::Row.new ["test", "asdf-asd"]
    r1.table = t
    expected = {test: Workbook::Cell.new("test"), asdfasd: Workbook::Cell.new("asdf-asd")}
    assert_equal(expected, r1.to_hash)
    date = DateTime.now
    r2 = Workbook::Row.new [2, date]
    r2.table = t
    expected = {test: Workbook::Cell.new(2), asdfasd: Workbook::Cell.new(date)}
    assert_equal(expected, r2.to_hash)
    assert_equal(date, r2[:asdfasd].value)
    assert_equal(date, r2[1].value)
  end

  def test_to_hash_with_values
    t = Workbook::Table.new
    r1 = Workbook::Row.new ["test", "asdf-asd"]
    r1.table = t
    expected = {test: "test", asdfasd: "asdf-asd"}
    assert_equal(expected, r1.to_hash_with_values)
    date = DateTime.now
    r2 = Workbook::Row.new [2, date]
    r2.table = t
    expected = {test: 2, asdfasd: date}
    assert_equal(expected, r2.to_hash_with_values)
    r3 = Workbook::Row.new [4]
    r3.table = t
    expected = {test: 4, asdfasd: nil}
    assert_equal(expected, r3.to_hash_with_values)
  end

  def test_to_hash_cache
    t = Workbook::Table.new
    t << ["test", "asdf-asd"]
    t << [1, 2]
    r = t.last
    assert_equal(1, r[:test].value)
    t.last[0].value = 3
    assert_equal(3, r[:test].value)
    assert_equal(3, r[:test].value)
    t.last[:test] = nil
    assert_nil(r[:test].value)
    r[:test] = 5
    assert_equal(5, r[:test].value)
  end

  def test_compare
    r1 = Workbook::Row.new ["test", "asdf-asd"]
    r2 = Workbook::Row.new [nil, "asdf-asd"]
    assert_equal(-1, r1 <=> r2)
    r1 = Workbook::Row.new [1, "asdf-asd"]
    r2 = Workbook::Row.new ["test", "asdf-asd"]
    assert_equal(-1, r1 <=> r2)
    r1 = Workbook::Row.new [nil, "asdf-asd"]
    r2 = Workbook::Row.new [Time.now, "asdf-asd"]
    assert_equal(1, r1 <=> r2)
    r1 = Workbook::Row.new [2, 3]
    r2 = Workbook::Row.new [2, nil]
    assert_equal(-1, r1 <=> r2)
    r1 = Workbook::Row.new [3, 0]
    r2 = Workbook::Row.new [2, 100000]
    assert_equal(1, r1 <=> r2)
    r1 = Workbook::Row.new [-10, 3]
    r2 = Workbook::Row.new [nil, 5]
    assert_equal(-1, r1 <=> r2)
  end

  def test_find_cells_by_background_color
    r = Workbook::Row.new ["test", "asdf-asd"]
    assert_equal([], r.find_cells_by_background_color)
    f = Workbook::Format.new
    f[:background_color] = "#ff00ff"
    r.first.format = f
    assert_equal([:test], r.find_cells_by_background_color)
    assert_equal([], r.find_cells_by_background_color("#ff0000"))
  end

  def test_to_s
    r1 = Workbook::Row.new ["test", "asdf-asd"]
    assert_equal("test,asdf-asd\n", r1.to_csv)
  end

  def test_clone
    b = Workbook::Book.new
    table = b.sheet.table
    table << Workbook::Row.new(["a", "b"])
    row = Workbook::Row.new(["1", "2"])
    table << row
    table << row
    row[1] = Workbook::Cell.new(3)
    table << table[1].clone
    table.last[1].value = 5
    assert_equal("a,b\n1,3\n1,3\n1,5\n", table.to_csv)
  end

  def test_clone_has_no_table
    b = Workbook::Book.new
    table = b.sheet.table
    table << Workbook::Row.new(["a", "b"])
    table << Workbook::Row.new([1, 2])
    row = table[1].clone
    assert_nil(row[:a])
    assert_nil(row[:b])
    assert_equal(1, row[0].value)
    assert_equal(2, row[1].value)
  end

  def test_push
    b = Workbook::Book.new
    table = b.sheet.table
    table << Workbook::Row.new(["a", "b"])
    table << Workbook::Row.new([1, 2])
    assert_equal(1, table[1][:a].value)
    assert_equal(2, table[1][:b].value)
    b = Workbook::Book.new
    table = b.sheet.table
    table.push Workbook::Row.new(["a", "b"])
    table.push Workbook::Row.new([1, 2])
    assert_equal(1, table[1][:a].value)
    assert_equal(2, table[1][:b].value)
  end

  def test_assign
    b = Workbook::Book.new
    table = b.sheet.table
    table.push Workbook::Row.new(["a", "b"])
    table[1] = Workbook::Row.new([1, 2])
    assert_equal(1, table[1][:a].value)
    assert_equal(2, table[1][:b].value)

    b = Workbook::Book.new
    table = b.sheet.table
    table.push Workbook::Row.new(["a", "b"])
    table[1] = [1, 2]
    assert_equal(1, table[1][:a].value)
    assert_equal(2, table[1][:b].value)
  end

  def test_preservation_of_format_on_assign
    row = Workbook::Row.new([1, 2])
    cellformat = row.first.format
    cellformat["background"] = "#f00"
    row[0] = 3
    assert_equal(3, row[0].value)
    assert_equal("#f00", row[0].format["background"])
  end

  def test_find_by_string
    b = Workbook::Book.new
    table = b.sheet.table
    table << Workbook::Row.new(["a", "b"])
    row = Workbook::Row.new([], table)
    row[1] = 12
    assert_equal(12, table.last["b"])
    assert_nil(table.last["a"])
  end

  def test_find_by_column_string
    b = Workbook::Book.new
    table = b.sheet.table
    table << Workbook::Row.new(["b", "a"])
    row = Workbook::Row.new([], table)
    row[1] = 12
    assert_equal(12, table.last["B"])
    assert_nil(table.last["A"])
  end

  def test_row_hash_index_string_assignment
    b = Workbook::Book.new
    table = b.sheet.table
    table << Workbook::Row.new(["a", "b", "d"])
    row = Workbook::Row.new([], table)
    row[1] = 12
    assert_equal(12, table.last.last.value)
    row[:b] = 15
    assert_equal(15, table.last.last.value)
    row["b"] = 18
    assert_equal(18, table.last.last.value)
    row["C"] = 2
    assert_equal(2, table.last[2].value)
  end

  def test_trim!
    a = Workbook::Row.new
    a[0] = 1
    a[1] = 2
    a[2] = nil
    b = Workbook::Row.new
    b[0] = 1
    b[1] = 2
    a.trim!
    assert_equal(b, a)
    a = Workbook::Row.new
    a[0] = nil
    a[1] = 2
    a[2] = nil
    b = Workbook::Row.new
    b[0] = nil
    b[1] = 2
    a.trim!
    assert_equal(b, a)
    a = Workbook::Row.new
    a[0] = 1
    a[1] = 2
    a[2] = nil
    b = Workbook::Row.new
    b[0] = 1
    b[1] = 2
    b[2] = nil
    a.trim!(3)
    assert_equal(b, a)
    a = Workbook::Row.new
    a[0] = 1
    a[1] = 2
    a[2] = nil
    b = Workbook::Row.new
    b[0] = 1
    b[1] = 2
    b[2] = nil
    b[3] = nil
    b[4] = nil
    b[5] = nil
    a.trim!(6)
    assert_equal(b, a)
    a = Workbook::Row.new
    a[0] = 1
    a[1] = 2
    a[2] = 3
    b = Workbook::Row.new
    b[0] = 1
    a.trim!(1)
    assert_equal(b, a)
  end

  def test_trim
    a = Workbook::Row.new
    a[0] = nil
    a[1] = 2
    a[2] = nil
    b = Workbook::Row.new
    b[0] = nil
    b[1] = 2
    b[2] = nil
    c = Workbook::Row.new
    c[0] = nil
    c[1] = 2
    d = a.trim
    assert_equal(b, a)
    assert_equal(c, d)
  end

  def test_add
    a = Workbook::Row.new
    a << 1
    a << 2
    a << "asdf"
    a << 2.2
    a.push(5)
    assert_equal(1, a[0].value)
    assert_equal(2, a[1].value)
    assert_equal("asdf", a[2].value)
    assert_equal(2.2, a[3].value)
    assert_equal(5, a[4].value)
  end

  def test_plus
    header = Workbook::Row.new([:a, :b])
    a = Workbook::Row.new
    table = Workbook::Table.new
    table << header
    table << a
    assert_equal(table, a.table)
    assert_equal(Workbook::Row, (a + [1, 1]).class)
    assert_equal([1, 1], (a + [1, 1]).to_a)
    assert_equal(Workbook::Cell, (a + [1, 1])[0].class)
    a += [1, 1]
    assert_equal([1, 1], a.to_a)
    assert_equal(Workbook::Row, a.class)
    assert_nil(a.table)
    assert_equal(Workbook::Cell, a[0].class)
  end

  def test_concat
    header = Workbook::Row.new([:a, :b])
    a = Workbook::Row.new
    table = Workbook::Table.new
    table << header
    table << a

    a.concat [1, 1]

    assert_equal([1, 1], a.to_a)
    assert_equal(Workbook::Row, a.class)
    assert_equal(table, a.table)
    assert_equal(Workbook::Cell, a[0].class)
  end
end
