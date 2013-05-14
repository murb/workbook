# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

class TestTemplate < Test::Unit::TestCase

  def test_initalize
    t = Workbook::Template.new
    assert_equal(true,(t.methods.include?(:add_raw) or t.methods.include?("add_raw")))
    assert_equal(true,(t.methods.include?(:has_raw_for?) or t.methods.include?("has_raw_for?")))
    assert_equal(true,(t.methods.include?(:raws) or t.methods.include?("raws")))
  end

  def test_add_raw_and_has_raw_for
    t = Workbook::Template.new
    t.add_raw "asdfsadf"
    assert_equal(false,t.has_raw_for?(Integer))
    assert_equal(true,t.has_raw_for?(String))
  end
  def test_raws
    t = Workbook::Template.new
    t.add_raw "asdfsadf"
    assert_equal({String=>"asdfsadf"}, t.raws)
  end

end
