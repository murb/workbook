# frozen_string_literal: true

require File.join(File.dirname(__FILE__), "helper")
module Readers
  class TestCsvWriter < Minitest::Test
    def test_to_csv
      w = Workbook::Book.new
      w.import File.join(File.dirname(__FILE__), "artifacts/simple_csv.csv")
      # reads
      #       a,b,c,d
      #       1,2,3,4
      #       5,3,2,1
      #       "asdf",123,12,2001-02-02
      #
      assert_equal("untitled document.csv", w.sheet.table.write_to_csv)
      csv_result = File.read("untitled document.csv").split("\n")
      csv_original = File.read(File.join(File.dirname(__FILE__), "artifacts/simple_csv.csv")).split("\n")
      assert_equal(csv_original[0], csv_result[0])
      assert_equal(csv_original[1], csv_result[1])
      assert_equal(csv_original[2], csv_result[2])
      assert_equal(csv_original[3], csv_result[3])
    end
  end
end
