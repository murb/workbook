# frozen_string_literal: true

require File.join(File.dirname(__FILE__), "helper")
module Readers
  class TestXlsShared < Minitest::Test
    def test_xls_number_to_time
      w = Workbook::Book.new
      assert_equal(DateTime.new(2011, 11, 15), w.xls_number_to_time(40862))
    end
  end
end
