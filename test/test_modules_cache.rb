# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

class TestToCacheClass
  include Workbook::Modules::Cache
  def basic_fetch
    fetch_cache(:a){
      sleep(0.5)
      "return"
    }
  end
  def expiring_fetch
    fetch_cache(:a,Time.now+0.6){
      sleep(0.5)
      "return"
    }
  end
end

module Modules
  class TestTableDiffSort < Minitest::Test
    def test_basic_fetch
      c = TestToCacheClass.new
      c.debug_cache = false
      start_time = Time.now
      c.basic_fetch
      diff_time = Time.now-start_time
      assert_equal(true,(diff_time > 0.5 and diff_time < 0.51))
      c.basic_fetch
      diff_time = Time.now-start_time
      assert_equal(true,(diff_time > 0.5 and diff_time < 0.51))
      c.basic_fetch
      diff_time = Time.now-start_time
      assert_equal(true,(diff_time > 0.5 and diff_time < 0.51))
    end
    def test_basic_fetch_invalidate_cache!
      c = TestToCacheClass.new
      start_time = Time.now
      c.basic_fetch
      diff_time = Time.now-start_time
      assert_equal(true,(diff_time > 0.5 and diff_time < 0.51))
      c.basic_fetch
      diff_time = Time.now-start_time
      assert_equal(true,(diff_time > 0.5 and diff_time < 0.51))
      c.invalidate_cache!
      c.basic_fetch
      diff_time = Time.now-start_time
      assert_equal(true,(diff_time > 1 and diff_time < 1.01))
    end
    def test_basic_fetch_invalidate_by_expiration
      c = TestToCacheClass.new
      start_time = Time.now
      c.expiring_fetch
      diff_time = Time.now-start_time
      assert_equal(true,(diff_time > 0.5 and diff_time < 0.51))
      c.expiring_fetch
      diff_time = Time.now-start_time
      assert_equal(true,(diff_time > 0.5 and diff_time < 0.51))
      c.expiring_fetch
      diff_time = Time.now-start_time
      assert_equal(true,(diff_time > 0.5 and diff_time < 0.51))
      sleep(0.5)
      diff_time = Time.now-start_time
      assert_equal(true,(diff_time > 1 and diff_time < 1.01))
    end

  end
end
