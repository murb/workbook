# frozen_string_literal: true

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
    def test_different_types_xlsx
      w = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/excel_different_types.xlsx')
      t = w.sheet.table
      assert_equal("ls",t["D4"].value)
      assert_equal(true,t["C3"].value)
      assert_equal("c",t["C1"].value)
      assert_equal(222,t["B3"].value)
      assert_equal(4.23,t["C2"].value)
      assert(t["A4"].value.is_a?(Date))
      assert((DateTime.new(2012,1,22,11)-t["A4"].value) < 0.00001)
      assert_equal(42000,t["B4"].value)
      assert_equal(42000.22323,t["D2"].value)
      assert(t["A2"].value.is_a?(Date))
      assert((Date.new(2012,2,22)-t["A2"].value) < 0.00001)
      assert((Date.new(2014,12,27)-t["B2"].value) < 0.00001)
      assert_equal(false,t["E2"].value)
      assert_equal(true,t["E3"].value)
      assert_equal(true,t["E4"].value)
    end
    def test_skipping_cells
      w = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/skippingcells.xlsx')
      t = w.sheet.table
      assert_equal("a,b,c,d,e,f,g,h,i,j,k,l,m,n\n1,,,,,,,,,,,,,\n,2,,,,,,,,,,,,\n,,3,,,,,,,,,,,\n,,,4,,,,,,,,,,\n,,,,5,,,,,,,,,\n,,,,,6,,,,,,,,\n,,,,,,7,,,,,,,\n,,,,,,,8,,,,,,\n,,,,,,,,9,,,,,\n,,,,,,,,,10,,,,\n,,,,,,,,,,11,,,\n,,,,,,,,,,,12,,\n,,,,,,,,,,,,13,\n,,,,,,,,,,,,,14\na,,,,,,,,,,,,,\n,b,,,,,,,,,,,,\n,,c,,,,,,,,,,,\n,,,d,,,,,,,,,,\n,,,,e,,,,,,,,,\n,,,,,f,,,,,,,,\n,,,,,,g,,,,,,,\n,,,,,,,h,,,,,,\n,,,,,,,,i,,,,,\n,,,,,,,,,j,,,,\n,,,,,,,,,,k,,,\n,,,,,,,,,,,l,,\n,,,,,,,,,,,,m,\n,,,,,,,,,,,,,n\n", t.to_csv)
    end
    def test_bit_table_xlsx
      b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/bigtable.xlsx')
      assert_equal(553, b.sheet.table.count)
    end
    def test_xlsx_with_empty_start
      b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/xlsx_with_empty_start.xlsx')
      t = b.sheet.table
      assert_nil(t["A3"].value)
    end
    def test_parse_shared_string_file
      file_contents = "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3095\" uniqueCount=\"1241\"><si><t>Nummer</t></si><si><t>Locatie</t></si><si><r><t>ZR 1</t></r><r><rPr><vertAlign val=\"superscript\"/><sz val=\"11\"/><rFont val=\"Calibri\"/><scheme val=\"minor\"/></rPr><t>e</t></r><r><rPr><sz val=\"11\"/><rFont val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></rPr><t xml:space=\"preserve\"> etage</t></r></si><si><t>Kaas</t></si></sst>"
      result = Workbook::Book.new.parse_shared_string_file(file_contents)
      assert_equal(["Nummer", "Locatie", "ZR 1e etage", "Kaas"],result)
    end
  end
end
