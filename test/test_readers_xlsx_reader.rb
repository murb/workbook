# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')
module Readers
  class TestXlsxReader < Minitest::Test
    def test_xlsx_open
      w = Workbook::Book.new
      w.import File.join(File.dirname(__FILE__), 'artifacts/book_with_tabs_and_colours.xlsx')
      assert_equal([:a, :b, :c, :d, :e],w.sheet.table.header.to_symbols)
      assert_equal([:anders,:dit],w[1].table.header.to_symbols)
      assert_equal(90588,w.sheet.table[2][:b].value)
      assert_equal(DateTime.new(2011,11,15),w.sheet.table[3][:d].value)
      # assert_equal("#CCFFCC",w.sheet.table[3][:c].format[:background_color]) #colour compatibility turned off for now...
      #assert_equal(8,w.sheet.table.first[:b].format[:width].round)
      # assert_equal(4,w.sheet.table.first[:a].format[:width].round)
      # assert_equal(25,w.sheet.table.first[:c].format[:width].round)
    end
    def test_open_native_xlsx
      w = Workbook::Book.new
      w.import File.join(File.dirname(__FILE__), 'artifacts/native_xlsx.xlsx')
      assert_equal([:datum_gemeld, :adm_gereed, :callnr],w.last.table.header.to_symbols)
      assert_equal("Callnr.",w.sheet.table[0][:callnr].value)
      assert_equal("2475617.00",w.sheet.table[3][:callnr].value)
      assert_equal("2012-12-03T12:30:00+00:00",w.sheet.table[7][:datum_gemeld].value.to_s)
      assert_equal("2012-12-03T09:4",w.sheet.table[6][:datum_gemeld].value.to_s[0..14])
    end
    def test_ms_formatting_to_strftime
      w = Workbook::Book.new
      assert_nil(w.ms_formatting_to_strftime(nil));
      assert_nil(w.ms_formatting_to_strftime(""));
    end
    def test_open_integer_xlsx
      w = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/integer_test.xlsx')
      assert_equal("2",w.sheet.table[1][1].value.to_s)
      assert_equal(2,w.sheet.table[1][1].value)
    end
    def test_bit_table_xlsx
      b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/bigtable.xlsx')
      assert_equal(553, b.sheet.table.count)
    end
  end
end
