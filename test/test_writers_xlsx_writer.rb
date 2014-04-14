# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

module Writers
  class TestXlsxWriter < Test::Unit::TestCase
    def test_empty_to_xlsx
      b = Workbook::Book.new [['a','b','c'],[1,2,3],[3,2,3]]
      b.to_xlsx.is_a? RubyXL::Workbook
      assert_equal('untitled document.xlsx', b.write_to_xlsx)

    end

    # def test_to_xlsx
  #     b = Workbook::Book.new [['a','b','c'],[1,2,3],[3,2,3]]
  #     raw = RubyXL::Parser.parse File.join(File.dirname(__FILE__), 'artifacts/simple_sheet.xlsx')
  #     t = Workbook::Template.new
  #     t.add_raw raw
  #     b.template = t
  #     assert_equal(true, b.to_xlsx.is_a?(RubyXL::Workbook))
  #
  #     assert_equal('untitled document.xlsx', b.write_to_xlsx)
  #   end

    def test_roundtrip
      b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/simple_sheet.xlsx')
      assert_equal(14,b[0][0]["A2"])
      assert_equal(DateTime.new(2011,11,15),b[0][0]["D3"].value)
      # puts b.sheet.table.to_csv
      filename = b.write_to_xlsx
      b = Workbook::Book.open filename
      assert_equal(14,b[0][0]["A2"])
      # assert_equal(DateTime.new(2011,11,15),b[0][0]["D3"].value) TODO: Dates don't work with RubyXL
    end
    def test_roundtrip_with_modification
      b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/simple_sheet.xlsx')
      b[0][0]["A2"]= 12
      assert_equal(DateTime.new(2011,11,15),b[0][0]["D3"].value)
      # puts b.sheet.table.to_csv
      filename = b.write_to_xlsx
      b = Workbook::Book.open filename
      assert_equal(12,b[0][0]["A2"].value)
      # assert_equal(DateTime.new(2011,11,15),b[0][0]["D3"].value) TODO: Dates don't work with RubyXL
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
      assert_equal(nil, t[4])
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
      t.pop(2)
      filename = b.write_to_xlsx
      b = Workbook::Book.open filename
      t = b.sheet.table
      assert_equal(33, t[3].first.value)
      assert_equal(nil, t[4])
      assert_equal(15, t[2].first.value)

    end
    # def test_pop_bigtable
    #   b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/bigtable.xlsx')
    #   # a  b  c  d  e
    #   # 14  90589  a  19 apr 12  23 apr 12
    #   # 15  90588  b  15 nov 11  16 jul 12
    #   # 25  90463  c  15 nov 11  17 nov 11
    #   # 33  90490  d  13 mrt 12  15 mrt 12
    #   t = b.sheet.table
    #   assert_equal(554, t.count)
    #   t.pop(300) #delete last two row
    #   assert_equal(254, t.trim.count)
    #   filename = b.write_to_xlsx
    #   b = Workbook::Book.open filename
    #   t = b.sheet.table
    #   assert_equal(254, t.trim.count)
    #
    #
    # end
    def test_cloning_roundtrip
      b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/book_with_tabs_and_colours.xlsx')
      b.sheet.table << b.sheet.table[2]
      assert_equal(90588,b.sheet.table[5][:b].value)
      assert_equal("#FFFF00",b.sheet.table[5][:c].format[:background_color])
      filename = b.write_to_xls
      b = Workbook::Book.open filename
      assert_equal(90588,b.sheet.table[5][:b].value)
      assert_equal("#FF00FF",b.sheet.table[5][:c].format[:background_color])
    end
 #    def test_parse_font_family
 #      b = Workbook::Book.new
 #      assert_equal(:none,b.parse_font_family({:font_family=>"asdfsdf"}))
 #      assert_equal(:swiss,b.parse_font_family({:font_family=>"ArIAL"}))
 #      assert_equal(:swiss,b.parse_font_family({:font_family=>:swiss}))
 #      assert_equal(:roman,b.parse_font_family({:font_family=>"Times"}))
 #      assert_equal(:roman,b.parse_font_family({:font_family=>"roman"}))
 #    end
 #
 #    def test_init_spreadsheet_template
 #      b = Workbook::Book.new
 #      b.init_spreadsheet_template
 #      assert_equal(Spreadsheet::Workbook,b.xls_template.class)
 #    end
 #
 #    def test_xls_sheet
 #      b = Workbook::Book.new
 #      b.init_spreadsheet_template
 #      assert_equal(Spreadsheet::Worksheet,b.xls_sheet(100).class)
 #    end
 #    def test_strftime_to_ms_format_nil
 #      assert_equal(nil, Workbook::Book.new.strftime_to_ms_format(nil))
 #    end
 #    def test_xls_sheet_writer
 #      b = Workbook::Book.new
 #      b << Workbook::Sheet.new
 #      b << Workbook::Sheet.new
 #      b[0].name = "A"
 #      b[1].name = "B"
 #      b[2].name = "C"
 #      assert_equal(["A","B","C"], b.collect{|a| a.name})
 #      filename = b.write_to_xls
 #      b = Workbook::Book.open filename
 #      assert_equal(["A","B","C"], b.collect{|a| a.name})
 #    end
 #    def test_removal_of_sheets_in_excel_when_using_template
 #      b = Workbook::Book.open File.join(File.dirname(__FILE__), 'artifacts/simple_sheet_many_sheets.xls')
 #      assert_equal(10, b.count)
 #      b.pop(4)
 #      assert_equal(6, b.count)
 #      filename = b.write_to_xls
 #      b = Workbook::Book.open filename
 #      assert_equal(6, b.count)
 #
 #    end
  end
end
