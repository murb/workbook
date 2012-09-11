require File.join(File.dirname(__FILE__), 'helper')

module Writers
  class TestXlsWriter < Test::Unit::TestCase
    def test_to_xls
      b = Workbook::Book.new [['a','b','c'],[1,2,3],[3,2,3]]
      raw = Spreadsheet.open('test/artifacts/simple_sheet.xls')
      t = Workbook::Template.new
      t.add_raw raw
      b.template = t
      assert_equal(true, b.to_xls.is_a?(Spreadsheet::Workbook))
      
      assert_equal('untitled document.xls', b.write_to_xls)
    end
    
    def test_roundtrip
      b = Workbook::Book.open('test/artifacts/simple_sheet.xls')
      assert_equal(3.85546875,b.sheet.table.first[:vestiging_id].format[:width])
      filename = b.write_to_xls
      b = Workbook::Book.open filename
      assert_equal(3.85546875,b.sheet.table.first[:vestiging_id].format[:width])
      
    end
    
    def test_init_spreadsheet_template
      b = Workbook::Book.new
      b.init_spreadsheet_template
      assert_equal(Spreadsheet::Workbook,b.xls_template.class)
    end
    
    def test_xls_sheet
      b = Workbook::Book.new
      b.init_spreadsheet_template
      assert_equal(Spreadsheet::Worksheet,b.xls_sheet(100).class)
    end
  end
end