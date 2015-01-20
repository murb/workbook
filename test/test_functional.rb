# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

class TestFunctional < Minitest::Test
  def test_chapter_Initializing
    b = Workbook::Book.new
    assert_equal(Workbook::Book, b.class)
    s = b.sheet
    assert_equal(Workbook::Sheet, s.class)
    t = s.table
    assert_equal(Workbook::Table, t.class)
    s = b.sheet[0] = Workbook::Sheet.new([['a','b'],[1,2],[3,4],[5,6]])
    assert_equal(Workbook::Sheet, s.class)

    t = s.table
    assert_equal(Workbook::Table, t.class)
    assert_equal(Workbook::Row, t.first.class)
    assert_equal(Workbook::Cell, t.first.first.class)
    assert_equal(true, t.header.header?)
    assert_equal(false, t.last.header?)
    assert_equal(2,t[1][:b].value)
  end

  def test_chapter_Sorting
    b = Workbook::Book.new
    s = b.sheet[0] = Workbook::Sheet.new([['a','b'],[1,2],[3,4],[5,6]])
    t = s.table
    #t.sort_by {|r| r[:b]}
    #p t.inspect
  end
end
