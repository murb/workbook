# frozen_string_literal: true

require File.join(File.dirname(__FILE__), "helper")

class TestModulesCell < Minitest::Test
  def test_init
    w = Workbook::Cell.new nil
    assert_nil(w.value)
    w = Workbook::Cell.new "asdf"
    assert_equal("asdf", w.value)

    assert_raises(ArgumentError) { w = Workbook::Cell.new Workbook::Row }

    t = Time.now
    w = Workbook::Cell.new t
    assert_equal(t, w.value)
  end

  def test_value
    w = Workbook::Cell.new nil
    assert_nil(w.value)
    w.value = "asdf"
    assert_equal("asdf", w.value)
    w.value = Date.new
    assert_equal(Date.new, w.value)
    w.value = 1
    assert_equal(1, w.value)
    assert(["Integer", "Fixnum"].include?(w.value.class.to_s))
    w.value = 1.0
    assert_equal(1.0, w.value)
    assert_equal(Float, w.value.class)
  end

  def test_importance_of_class
    a = Workbook::Cell.new
    assert_equal(4, a.importance_of_class("a"))
    assert_equal(5, a.importance_of_class(:a))
  end

  def test_comp
    a = Workbook::Cell.new 1
    b = Workbook::Cell.new 2
    assert_equal(-1, a <=> b)
    a = Workbook::Cell.new "c"
    b = Workbook::Cell.new "bsdf"
    assert_equal(1, a <=> b)
    a = Workbook::Cell.new "c"
    b = Workbook::Cell.new "c"
    assert_equal(0, a <=> b)
    a = Workbook::Cell.new true
    b = Workbook::Cell.new false
    assert_equal(-1, a <=> b)
    a = Workbook::Cell.new "true"
    b = Workbook::Cell.new "false"
    assert_equal(1, a <=> b)
    a = Workbook::Cell.new 1
    b = Workbook::Cell.new "a"
    assert_equal(-1, a <=> b)
    a = Workbook::Cell.new nil
    b = Workbook::Cell.new "a"
    assert_equal(1, a <=> b)
    a = Workbook::Cell.new nil
    b = Workbook::Cell.new nil
    assert_equal(0, a <=> b)
  end

  def test_cloning_as_expected?
    a = Workbook::Cell.new 1
    a.format = Workbook::Format.new({value: 1})
    b = a.clone
    b.value = 2
    b.format[:value] = 2
    assert_equal(1, a.value)
    assert_equal(2, b.value)
    assert_equal(2, a.format[:value])
    assert_equal(2, b.format[:value])
  end

  def test_to_sym
    examples = {
      "A - B" => :a_b,
      "A-B" => :ab,
      "A - c (B123)" => :a_c_b123,
      "A - c (B123)!" => :a_c_b123!,
      "A-B?" => :ab?,
      "A-B!" => :ab!,
      "éåšžÌ?" => :easzi?,
      1 => :num1,
      1.0 => :num1_0,
      "test   " => :test,
      "test " => :test,

    }
    examples.each do |k, v|
      assert_equal(v, Workbook::Cell.new(k).to_sym)
    end
  end

  def test_nil
    c = Workbook::Cell.new nil
    assert_nil(c)
  end

  def test_colspan_rowspan
    c = Workbook::Cell.new
    c.colspan = 1
    c.rowspan = 1
    assert_nil(c.colspan)
    assert_nil(c.rowspan)
    c.colspan = nil
    c.rowspan = ""
    assert_nil(c.colspan)
    assert_nil(c.rowspan)
    c.colspan = 3
    c.rowspan = "4"
    assert_equal(3, c.colspan)
    c.rowspan = 0
    assert_nil(c.rowspan)
    assert_equal(3, c.colspan)
    c.colspan = 0
    c.rowspan = 3
    assert_equal(3, c.rowspan)
    assert_nil(c.colspan)
  end

  def test_cell_type
    {1 => :integer, 3.2 => :float, true => :boolean, "asdf" => :string}.each do |k, v|
      c = Workbook::Cell.new(k)
      assert_equal(v, c.cell_type)
    end
  end

  def test_index
    t = Workbook::Table.new [[:a, :b, :c], [1, 2, 3], [4, 5, 6]]
    assert_equal(2, t[2][2].index)
  end

  def test_key
    t = Workbook::Table.new [[:a, :b, :c], [1, 2, 3], [4, 5, 6]]
    assert_equal(:c, t[2][2].key)
    t = Workbook::Table.new [[:d, nil, nil], [:a, :b, :c], [1, 2, 3], [4, 5, 6]]
    assert_nil(t[2][2].key)
    assert_equal(:d, t[2][0].key)
    t.header = t[1]
    assert_equal(:a, t[2][0].key)
  end
end
