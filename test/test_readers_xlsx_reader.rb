require File.join(File.dirname(__FILE__), 'helper')
module Readers
  class TestXlsxWriter < Test::Unit::TestCase
    def test_open
      w = Workbook::Book.new
      w.open 'test/artifacts/book_with_tabs_and_colours.xlsx'
      assert_equal([:a, :b, :c, :d, :e],w.sheet.table.header.to_symbols)
      assert_equal(90588,w.sheet.table[2][:b].value)
      assert_equal(DateTime.new(2011,11,15),w.sheet.table[3][:d].value)
     # assert_equal("#CCFFCC",w.sheet.table[3][:c].format[:background_color]) #colour compatibility turned off for now...
      assert_equal(8,w.sheet.table.first[:b].format[:width].round)
      assert_equal(4,w.sheet.table.first[:a].format[:width].round)
      assert_equal(25,w.sheet.table.first[:c].format[:width].round)
      y w
    end
  end
end