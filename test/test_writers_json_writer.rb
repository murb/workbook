# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

module Writers
  class TestJsonWriter < Test::Unit::TestCase
    def test_to_array_of_hashes_with_values
      assert_equal([],Workbook::Table.new.to_array_of_hashes_with_values)
      assert_equal([{:a=>"a",:b=>"b"}],Workbook::Table.new([["a","b"]]).to_array_of_hashes_with_values)
      assert_equal([{:a=>"a",:b=>"b"},{:a=>1,:b=>2},{:a=>Date.new(2012,1,1),:b=>nil}],
        Workbook::Table.new([["a","b"],[1,2],[Date.new(2012,1,1),nil]]).to_array_of_hashes_with_values)
    end
    def test_to_json
      assert_equal("[]",Workbook::Table.new.to_json)
      assert_equal("[{\"a\":\"a\",\"b\":\"b\"}]",Workbook::Table.new([["a","b"]]).to_json)
      assert_equal("[{\"a\":\"a\",\"b\":\"b\"},{\"a\":1,\"b\":2},{\"a\":\"2012-01-01\",\"b\":null}]",
        Workbook::Table.new([["a","b"],[1,2],[Date.new(2012,1,1),nil]]).to_json)
    end
  end
end
