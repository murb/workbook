# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')
class TestToCacheClass
  include Workbook::Modules::Cache

  def a
    fetch_cache(:a, Time.now-20000) {
      sleep(5)
      "het antwoord"
    }
  end
end

module Modules
  class TestTableDiffSort < Test::Unit::TestCase
    def test_basic_fetch
      c = TestToCacheClass.new
      puts c.a
      puts c.a
      puts c.a
    end

  end
end
