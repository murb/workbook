# frozen_string_literal: true

require File.join(File.dirname(__FILE__), "helper")

module Writers
  class TestJsonWriter < Minitest::Test
    def test_to_array_of_hashes_with_values
      assert_equal([], Workbook::Table.new.to_array_of_hashes_with_values)
      assert_equal([], Workbook::Table.new([["a", "b"]]).to_array_of_hashes_with_values)
      assert_equal([{a: 1, b: 2}, {a: Date.new(2012, 1, 1), b: nil}],
        Workbook::Table.new([["a", "b"], [1, 2], [Date.new(2012, 1, 1), nil]]).to_array_of_hashes_with_values)
    end

    def test_to_json
      assert_equal("[]", Workbook::Table.new.to_json)
      assert_equal("[]", Workbook::Table.new([["a", "b"]]).to_json)
      assert_equal("[{\"a\":1,\"b\":2},{\"a\":\"2012-01-01\",\"b\":null}]",
        Workbook::Table.new([["a", "b"], [1, 2], [Date.new(2012, 1, 1), nil]]).to_json)
    end

    def test_write_to_json
      filename = Workbook::Table.new([["a", "b"], [1, 2], [Date.new(2012, 1, 1), nil]]).write_to_json("json_test.json")
      json_written = File.read(filename)
      assert_equal("[{\"a\":1,\"b\":2},{\"a\":\"2012-01-01\",\"b\":null}]",json_written)

    end
  end
end
