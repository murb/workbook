# frozen_string_literal: true

require File.join(File.dirname(__FILE__), "helper")

class TestFormat < Minitest::Test
  def test_initialize
    f = Workbook::Format.new {}
    assert_equal({}, f)
    f = Workbook::Format.new({background: :red})
    assert_equal({background: :red}, f)
    f = Workbook::Format.new({background: :red})
    deet = Time.now
    assert_equal(false, f.has_raw_for?(Time))

    f.add_raw deet
    assert_equal(deet, f.raws[Time])
    assert_equal(true, f.has_raw_for?(Time))
  end

  def test_derived_type
    f = Workbook::Format.new {}
    assert_nil(f.derived_type)
    f = Workbook::Format.new numberformat: "%m/%e/%y h:%m"
    assert_equal(:time, f.derived_type)
    f = Workbook::Format.new numberformat: "%m/%e/%y"
    assert_equal(:date, f.derived_type)
  end

  def test_available_raws
    deet = Time.now
    f = Workbook::Format.new {}
    assert_equal([], f.available_raws)
    f.add_raw deet
    assert_equal([Time], f.available_raws)
  end

  def test_merge
    a = Workbook::Format.new({background: :red})
    b = Workbook::Format.new({background: :yellow, color: :green})
    result = b.clone.merge(a)
    assert_equal({background: :red, color: :green}, result)
    assert_equal(true, result.is_a?(Workbook::Format))
  end

  def test_remove_raw_on_merge
    a = Workbook::Format.new({background: :red})
    b = Workbook::Format.new({background: :yellow, color: :green})
    b.add_raw Date.new
    result = b.clone.merge(a)
    assert_equal({}, result.raws)
  end

  def test_has_background_color?
    a = Workbook::Format.new
    assert_equal(false, a.has_background_color?)
    a = Workbook::Format.new({background_color: "#ff0000"})
    assert_equal(true, a.has_background_color?)
    assert_equal(false, a.has_background_color?("#00ff00"))
    a = Workbook::Format.new({background_color: "#ffffff"})
    assert_equal(false, a.has_background_color?)
    a = Workbook::Format.new({background_color: "#FFFFFf"})
    assert_equal(false, a.has_background_color?)
  end

  def test_to_css
    a = Workbook::Format.new({background_color: "#ffffff"})
    assert_equal("background: #ffffff", a.to_css)
    a = Workbook::Format.new({background_color: "#fffdff", color: :red})
    assert_equal("background: #fffdff; color: red", a.to_css)
  end

  def test_parent_style
    a = Workbook::Format.new({background_color: "#ffffff"})
    b = Workbook::Format.new({color: "#000"})
    a.parent = b
    assert_equal(b, a.parent)
  end

  def test_parent_style_flattened_properties
    a = Workbook::Format.new({background_color: "#ffffff"})
    b = Workbook::Format.new({color: "#000"})
    a.parent = b
    assert_equal(Workbook::Format.new({color: "#000", background_color: "#ffffff"}), a.flattened)
    c = Workbook::Format.new({color: "#f00"})
    c.parent = a
    assert_equal(Workbook::Format.new({color: "#f00", background_color: "#ffffff"}), c.flattened)
  end

  def test_formats
    a = Workbook::Format.new({background_color: "#ffffff"})
    b = Workbook::Format.new({color: "#000"})
    c = Workbook::Format.new({color: "#f00"})
    a.parent = b
    c.parent = a
    assert_equal([b, a, c], c.formats)
  end

  def test_parent_style_to_css
    a = Workbook::Format.new({background_color: "#ffffff"})
    b = Workbook::Format.new({color: "#000"})
    c = Workbook::Format.new({color: "#f00"})
    a.parent = b
    c.parent = a
    assert_equal("background: #ffffff; color: #f00", c.to_css)
  end

  def test_all_names
    a = Workbook::Format.new({background_color: "#ffffff"}, "a")
    b = Workbook::Format.new({color: "#000"}, "b")
    c = Workbook::Format.new({color: "#f00"}, "c")
    a.parent = b
    c.parent = a
    assert_equal(["b", "a", "c"], c.all_names)
  end
end
