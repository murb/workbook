# frozen_string_literal: true

# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

module Writers
  class TestXlsxWriter < Minitest::Test
    def test_empty_to_xlsx
      b = Workbook::Book.new [['a','b','c'],[1,2,3],[3,2,3]]
      assert_equal(true, b.to_xlsx.is_a?(Axlsx::Package))
      dimensions = b.sheet.table.dimensions
      assert_equal('untitled document.xlsx', b.write_to_xlsx)
      b = Workbook::Book.open 'untitled document.xlsx'
      assert_equal(dimensions, b.sheet.table.dimensions)
    end

    def test_roundtrip
      b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/simple_sheet.xlsx')
      assert_equal(14,b[0][0]["A2"])
      assert_equal(DateTime.new(2011,11,15),b[0][0]["D3"].value)
      # puts b.sheet.table.to_csv
      filename = b.write_to_xlsx
      b = Workbook::Book.open filename
      assert_equal(14,b[0][0]["A2"].value)
      assert_equal(DateTime.new(2011,11,15),b[0][0]["D3"].value)
    end
    def test_roundtrip_with_modification
      b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/simple_sheet.xlsx')
      b[0][0]["A2"]= 12
      assert_equal(DateTime.new(2011,11,15),b[0][0]["D3"].value)
      filename = b.write_to_xlsx
      b = Workbook::Book.open filename
      assert_equal(12,b[0][0]["A2"].value)
      assert_equal(DateTime.new(2011,11,15),b[0][0]["D3"].value)
    end
    def test_delete_row
      b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/simple_sheet.xlsx')
      # a  b  c  d  e
      # 14  90589  a  19 apr 12  23 apr 12
      # 15  90588  b  15 nov 11  16 jul 12
      # 25  90463  c  15 nov 11  17 nov 11
      # 33  90490  d  13 mrt 12  15 mrt 12
      t = b.sheet.table
      assert_equal(33, t.last.first.value)
      t.delete_at(4) #delete last row
      filename = b.write_to_xlsx
      b = Workbook::Book.open filename
      t = b.sheet.table
      #TODO: NOT true delete... need to work on this...
      assert_equal(25, t[3].first.value)
      assert_nil(t[4])
    end
    def test_pop_row
      b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/simple_sheet.xlsx')
      # a  b  c  d  e
      # 14  90589  a  19 apr 12  23 apr 12
      # 15  90588  b  15 nov 11  16 jul 12
      # 25  90463  c  15 nov 11  17 nov 11
      # 33  90490  d  13 mrt 12  15 mrt 12
      t = b.sheet.table
      assert_equal(33, t.last.first.value)
      t.pop(1)
      filename = b.write_to_xlsx
      b = Workbook::Book.open filename
      t = b.sheet.table
      assert_equal(25, t[3].first.value)
      assert_nil(t[4])
      assert_equal(15, t[2].first.value)
    end
    def test_pop_bigtable
      b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/bigtable.xlsx')
      t = b.sheet.table
      assert_equal(553, t.count)
      t.pop(300)
      assert_equal(253, t.trim.count)
      filename = b.write_to_xlsx
      b = Workbook::Book.open filename
      t = b.sheet.table
      assert_equal(253, t.trim.count)
    end

    # Uncommented colour testing, this is broken since the switch to roo/axlsx
    def test_cloning_roundtrip
      b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/book_with_tabs_and_colours.xlsx')
      b.sheet.table << b.sheet.table[2]
      assert_equal(90588,b.sheet.table[5][:b].value)
      # assert_equal("#FFFF00",b.sheet.table[2][:c].format[:background_color])
      # assert_equal("#FFFF00",b.sheet.table[5][:c].format[:background_color])
      filename = b.write_to_xls
      b = Workbook::Book.open filename
      assert_equal(90588,b.sheet.table[5][:b].value)
      # assert_equal("#FF00FF",b.sheet.table[5][:c].format[:background_color])
    end

    def test_format_to_xlsx_format
      b = Workbook::Book.new
      xlsx_format = b.format_to_xlsx_format(Workbook::Format.new({font_weight: "bold", color: "#FF0000"}))
      assert_equal(true,xlsx_format[:b])
      assert_equal("FFFF0000",xlsx_format[:fg_color])
    end

    def test_formats_to_xlsx_format
      b = Workbook::Book.new
      b.template.set_default_formats!
      b.formats_to_xlsx_format
      raw_keys = b.template.create_or_find_format_by(:header).raws.keys
      assert((raw_keys.include?(Integer) or raw_keys.include?(Fixnum)))
    end

    def test_format_to_xlsx_integrated
      b = Workbook::Book.new [[:a,:b],[1,2],[3,4]]
      c2 = b.sheet.table[2][1]
      c2.format = Workbook::Format.new({font_weight: "bold", color: "#CC5500", font_style: :italic, text_decoration: :underline})
      # Can't test this for real yet... :/ but the examples here seem to work     b.write_to_xlsx("untitled document.xlsx")
      # c = Workbook::Book.open("untitled document.xlsx")
      # p c.inspect
    end

  end
end
