require 'test/helper'
module Readers
  class TestXlsWriter < Test::Unit::TestCase
    def test_open
      w = Workbook::Book.new
      w.open 'test/artifacts/book_with_tabs_and_colours.xls'
      assert_equal([:vestiging_id,:pirnr_ing,:ing_kantoornaam,:installatiedag_ispapparatuur,:mo_openingsdatum_ing],w.sheet.table.header.to_symbols)
      assert_equal(90588,w.sheet.table[2][:pirnr_ing].value)
      assert_equal("#CCFFCC",w.sheet.table[3][:ing_kantoornaam].format[:background_color])
    end
    
    def test_complex_types
      w = Workbook::Book.new
      w.open 'test/artifacts/complex_types.xls'
      assert_equal(Date.new(2011,11,15), w.sheet.table[2][3].value)
      assert_equal("http://murb.nl", w.sheet.table[3][2].value)
      assert_equal("sadfasdfsd", w.sheet.table[4][2].value)
      assert_equal(1.2, w.sheet.table[3][1].value)
    end
  end
end