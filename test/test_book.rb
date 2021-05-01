# frozen_string_literal: true

require File.join(File.dirname(__FILE__), "helper")

class TestWorkbook < Minitest::Test
  def test_sheets
    w = Workbook::Book.new nil
    w.push
    assert_equal(1, w.count)
  end

  def test_push
    w = Workbook::Book.new nil
    assert_equal(0, w.count)
    assert_equal([], w)

    w = Workbook::Book.new
    w.push
    assert_equal(w.first.class, Workbook::Sheet)

    w.push
    assert_equal(w, w.last.book)
    assert_equal(2, w.count)

    s = Workbook::Sheet.new([[1, 2], [3, 4]])
    w.push s
    assert_equal(w, w.last.book)

    assert_equal(w.last, s)

    w = Workbook::Book.new
    assert_equal(w.sheet.table.class, Workbook::Table)
  end

  def test_sheet
    w = Workbook::Book.new nil
    s = Workbook::Sheet.new [Workbook::Row.new(Workbook::Table.new)]
    assert_equal(w.sheet.class, Workbook::Sheet)
    refute_equal(w.sheet, s)
    w = Workbook::Book.new s
    assert_equal(w.sheet, s)
    w = Workbook::Book.new
    s = w.sheet
    assert_equal(Workbook::Sheet, s.class)
  end

  def test_template
    b = Workbook::Book.new
    raw = "asdf"
    assert_raises(ArgumentError) { b.template = raw }
    raw = Workbook::Template.new
    b.template = raw

    assert_equal(raw, b.template)
  end

  def test_file_extension
    b = Workbook::Book.new
    assert_equal("aaa", b.file_extension("aaa.aaa"))
    b = Workbook::Book.new
    assert_equal("xlsx", b.file_extension(File.join(File.dirname(__FILE__), "artifacts/book_with_tabs_and_colours.xlsx")))
    b = Workbook::Book.new
    assert_equal("xlsx", b.file_extension(File.new(File.join(File.dirname(__FILE__), "artifacts/book_with_tabs_and_colours.xlsx"))))
  end

  def test_parent_child
    b = Workbook::Book.new [[1, 2, 3], [1, 2, 3]]
    assert_equal(Workbook::Sheet, b.first.class)
    assert_equal(b, b.first.book)
    assert_equal(Workbook::Table, b.first.table.class)
    assert_equal(b, b.first.table.sheet.book)
    assert_equal(Workbook::Row, b.first.table.header.class)
    assert_equal(b, b.first.table.header.table.sheet.book)
  end

  def test_text_to_utf8
    f = File.open(File.join(File.dirname(__FILE__), "artifacts/excel_different_types.txt"), "r")
    t = f.read
    w = Workbook::Book.new
    t = w.text_to_utf8(t)
    assert_equal("a\tb\tc\td", t.split(/(\n|\r)/).first)
  end

  def test_read_bad_filetype
    assert_raises(ArgumentError) { Workbook::Book.read("test string here", :xls) }
    assert_raises(ArgumentError) { Workbook::Book.read("test string here", :ods) }
    assert_raises(ArgumentError) { Workbook::Book.read("test string here", :xlsx) }
  end

  def test_push_and_ltlt
    b = Workbook::Book.new [["a", "b"], [1, 2]]
    b.push Workbook::Sheet.new([["a", "b"], [2, 2]])
    b.push Workbook::Sheet.new([["a", "b"], [3, 2]])
    b << Workbook::Sheet.new([["a", "b"], [4, 2]])
    b.push Workbook::Sheet.new([["a", "b"], [5, 2]])
    b << Workbook::Sheet.new([["a", "b"], [6, 2]])
    b.push Workbook::Sheet.new([["a", "b"], [7, 2]])

    # puts b.index b.last
    7.times { |time| assert_equal(b, b[0].book) }
  end

  def test_removal_of_sheets_pop_and_delete_at_works_as_expected
    b = Workbook::Book.new [["a", "b"], [1, 2]]
    b.push Workbook::Sheet.new([["a", "b"], [2, 2]])
    b.push Workbook::Sheet.new([["a", "b"], [3, 2]])
    b << Workbook::Sheet.new([["a", "b"], [4, 2]])
    b.push Workbook::Sheet.new([["a", "b"], [5, 2]])
    b << Workbook::Sheet.new([["a", "b"], [6, 2]])
    b.push Workbook::Sheet.new([["a", "b"], [7, 2]])

    assert_equal(7, b.count)
    assert_equal(5, b[4][0][1][0].value)
    b.delete_at(4)
    assert_equal(6, b.count)
    assert_equal(6, b[4][0][1][0].value)
    b.pop(3)
    assert_equal(3, b.count)
  end

  def test_supported_mime_types
    assert_equal true, Workbook::SUPPORTED_MIME_TYPES.include?("text/csv")
  end
end
