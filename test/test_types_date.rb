# frozen_string_literal: true

# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

class TestTypesDate < Minitest::Test
  def test_init
    w = Workbook::Types::Date.new(2001,2,2)
    assert_equal(Workbook::Types::Date, w.class)
    assert_equal(Date.new(2001,2,2),w)
    assert_equal(Date.new(2001,2,2),w.value)
    assert_equal(true, w.is_a?(Date))
    assert_equal(true, w.is_a?(Workbook::Modules::Cell))

  end
end
