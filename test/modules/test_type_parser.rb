require File.join(File.dirname(__FILE__), '../helper')
module Modules
  class TestTypeParser < Test::Unit::TestCase
    def examples
      {"2312"=>2312,
       "12-12-2012"=>Date.new(2012,12,12),
       "12-12-2012 12:24"=>DateTime.new(2012,12,12,12,24),
       "1-11-2011"=>Date.new(2011,11,1),
       "jA"=>true,
       "n"=>false,
       "12 bomen"=>"12 bomen",
       "12 bomenasdfasdfsadf"=>"12 bomenasdfasdfsadf",
       ""=>nil,
       " "=>nil,
       "mailto:sadf@asdf.as"=>"sadf@asdf.as"
       }
    end
    
    def test_parse
      examples.each do |k,v|
        assert_equal(v,Workbook::Cell.new(k).parse({:detect_date=>true}))
      end
    end
    
    def test_custom_parse
      customparsers = [proc{|v| "#{v}2" }]
      examples.each do |k,v|
        c = Workbook::Cell.new(k)
        c.string_parsers = customparsers
        assert_not_equal(v,c.parse)
      end
      c = Workbook::Cell.new("233")
      c.string_parsers = customparsers
      assert_equal("2332",c.parse)
      c.value = "v"
      assert_equal("v2",c.parse)
    end
    
    def test_parse!
      r= Workbook::Row.new
      r[0] = Workbook::Cell.new "xls_cell"    
      r[0].parse!    
      assert_equal("xls_cell",r[0].value)
      r[1] = Workbook::Cell.new ""    
      r[1].parse!    
      assert_equal(nil,r[1].value)
      
    end
  end
end
