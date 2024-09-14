# frozen_string_literal: true

require File.join(File.dirname(__FILE__), "helper")
module Readers
  class TestOdsWriter < Minitest::Test
    def test_ods_open
      w = Workbook::Book.new
      w.import File.join(File.dirname(__FILE__), "artifacts/book_with_tabs_and_colours.ods")

      assert_equal([:a, :b, :c, :d, :e], w.sheet.table.header.to_symbols)
      assert_equal(90588, w.sheet.table[2][:b].value)
    end

    def test_styling
      w = Workbook::Book.new
      w.import File.join(File.dirname(__FILE__), "artifacts/book_with_tabs_and_colours.ods")
      assert_equal("#ffff99", w.sheet.table[3][:c].format[:background_color])
      assert_equal(true, w.sheet.table[0][:e].format.all_names.include?("Heading1"))
      # TODO: column styles
      # assert_equal(8.13671875,w.sheet.table.first[:b].format[:width])
      # assert_equal(3.85546875,w.sheet.table.first[:a].format[:width])
      # assert_equal(25.14453125,w.sheet.table.first[:c].format[:width])
    end

    def test_complex_types
      w = Workbook::Book.new
      w.import File.join(File.dirname(__FILE__), "artifacts/complex_types.ods")
      assert_equal(Date.new(2011, 11, 15), w.sheet.table[2][3].value)
      assert_equal("http://murb.nl", w.sheet.table[3][2].value)
      assert_equal("Sadfasdfsd > 2", w.sheet.table[4][2].value)
      assert_equal(1.2, w.sheet.table[3][1].value)
    end

    def test_currency
      w = Workbook::Book.new
      w.import File.join(File.dirname(__FILE__), "artifacts/currency_test.ods")

      assert_equal(1200, w.sheet.table["H2"].value)
      assert_equal(1200.4, w.sheet.table["H4"].value)
    end

    def test_excel_standardized_open
      w = Workbook::Book.new
      w.import File.join(File.dirname(__FILE__), "artifacts/excel_different_types.ods")
      assert_equal([:a, :b, :c, :d], w.sheet.table.header.to_symbols[0..3])
      assert_equal(Date.new(2012, 2, 22), w.sheet.table[1][:a].value)
      assert_equal("c", w.sheet.table[2][:a].value)
      assert_equal(DateTime.new(2012, 1, 22, 11), w.sheet.table[3][:a].value)
      assert_equal("42000", w.sheet.table[3][:b].value.to_s)
      assert_nil(w.sheet.table[2][:c].value)
    end

    def test_sheet_with_combined_cells
      w = Workbook::Book.new
      w.import File.join(File.dirname(__FILE__), "artifacts/sheet_with_combined_cells.ods")
      t = w.sheet.table
      assert_equal("14 90589", t[1][:a].value)
      assert_equal(Workbook::NilValue, t[1][:b].value.class)
      assert_equal(:covered, t[1][:b].value.reason)
      assert_equal(2, t[1][:a].colspan)
      assert_nil(t[1][:c].colspan)
      assert_equal(2, t["D3"].rowspan)
      assert_equal(2, t["D5"].rowspan)
      assert_equal(2, t["D5"].colspan)
    end

    def test_book_with_colspans
      w = Workbook::Book.new
      w.import File.join(File.dirname(__FILE__), "artifacts/book_with_colspans.ods")
      t = w.sheet.table
      assert_equal(2, t["B1"].colspan)
      assert_equal(2, t["D1"].colspan)
      assert_nil(t["D3"].value)
      assert_equal("g", t["A19"].value)
      assert_equal(0.03, t["D17"].value)
    end
  end
end
