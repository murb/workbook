# frozen_string_literal: true

require File.join(File.dirname(__FILE__), "helper")

module Writers
  class TestHtmlWriter < Minitest::Test
    def test_to_html
      # jruby and ruby's output differ a bit... both produce valid results though
      # match = Workbook::Book.new.to_html.match(/<table \/>/) ? true : false #jruby
      # puts Workbook::Book.new.to_html
      # match = (Workbook::Book.new.to_html.match(/<table><thead><\/thead><tbody><\/tbody><\/table>/) ? true : false) if match == false #ruby
      # assert_equal(true, match)
      html = Workbook::Book.new([["a", "b"], [1, 2], [3, 4]]).to_html
      match = /<table><\/table>/.match?(html) ? true : false
      assert_equal(false, match)
      match = /<td>1<\/td>/.match?(html) ? true : false
      assert_equal(true, match)
      match = /<th class=\"a\" data-key=\"a\">a<\/th>/.match?(html) ? true : false
      assert_equal(true, match)
    end

    def test_to_html_format_names
      b = Workbook::Book.new([["a", "b"], [1, 2], [3, 4]])
      c = b[0][0][0][0]
      c.format.name = "testname"
      c = b[0][0][1][0]
      c.format.name = "testname"
      html = b.to_html
      match = /<th class=\"testname a\" data-key=\"a\">a<\/th>/.match?(html) ? true : false
      assert_equal(true, match)
      match = /<td class=\"testname\">1<\/td>/.match?(html) ? true : false
      assert_equal(true, match)
    end

    def test_build_cell_options
      b = Workbook::Book.new([["a", "b"], [1, 2], [3, 4]])
      result = b.sheet.table.build_cell_options(b.sheet.table.first.first, {data: {a: "a"}})
      assert("a", result["data-a"])
    end

    def test_to_html_css
      b = Workbook::Book.new([["a", "b"], [1, 2], [3, 4]])
      c = b[0][0][0][0]
      c.format[:background] = "#f00"
      c = b[0][0][1][0]
      c.format[:background] = "#ff0"
      html = b.to_html
      match = /<th class=\"a\" data-key=\"a\">a<\/th>/.match?(html) ? true : false
      assert_equal(true, match)
      match = /<td>1<\/td>/.match?(html) ? true : false
      assert_equal(true, match)
      html = b.to_html({style_with_inline_css: true})
      match = /<th class=\"a\" data-key=\"a\" style=\"background: #f00\">a<\/th>/.match?(html) ? true : false
      assert_equal(true, match)
      match = /<td style="background: #ff0">1<\/td>/.match?(html) ? true : false
      assert_equal(true, match)
    end

    def test_sheet_and_table_names
      b = Workbook::Book.new([["a", "b"], [1, 2], [3, 4]])
      b.sheet.name = "Sheet name"
      b.sheet.table.name = "Table name"
      html = b.to_html
      assert_equal(true, (/<h1>Sheet name<\/h1>/.match?(html) ? true : false))
      assert_equal(true, (/<h2>Table name<\/h2>/.match?(html) ? true : false))
    end

    def test_col_and_rowspans
      w = Workbook::Book.new
      w.import File.join(File.dirname(__FILE__), "artifacts/sheet_with_combined_cells.ods")
      html = w.to_html
      assert_equal(true, (/rowspan="2">15 nov 11 15 nov 11/.match?(html) ? true : false))
      if RUBY_VERSION >= "1.9"
        assert_equal(true, (/colspan="2" rowspan="2">13 mrt 12 15 mrt 12 13 mrt 12 15 mrt 12/.match?(html) ? true : false))
        assert_equal(true, (/colspan="2">14 90589/.match?(html) ? true : false))
      end
    end
  end
end
