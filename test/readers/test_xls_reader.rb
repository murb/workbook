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
  end
end