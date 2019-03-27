# frozen_string_literal: true

# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

class TestTemplate < Minitest::Test

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
  def test_set_default_formats!
    t = Workbook::Template.new
    t.set_default_formats!
    assert_equal({font_weight: "bold"},t.formats[:header][:default])
  end
  def test_add_formats
    t = Workbook::Template.new
    t.add_format Workbook::Format.new({font:"Arial"})
    t.add_format Workbook::Format.new({font:"Times"})
    assert_equal(2,t.formats.keys.count)
    named_format = Workbook::Format.new({font:"Times"})
    named_format.name = 1
    t.add_format named_format
    assert_equal(2,t.formats.keys.count)
  end
end
