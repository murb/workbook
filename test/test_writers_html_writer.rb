# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')

module Writers
  class TestXlsWriter < Test::Unit::TestCase
    def test_to_html
      match = Workbook::Book.new.to_html.match(/<table><\/table>/) ? true : false
      assert_equal(true, match)
      html = Workbook::Book.new([['a','b'],[1,2],[3,4]]).to_html
      match = html.match(/<table><\/table>/) ? true : false
      assert_equal(false, match)
      match = html.match(/<td>1<\/td>/) ? true : false
      assert_equal(true, match)
      match = html.match(/<td>a<\/td>/) ? true : false
      assert_equal(true, match)
    end
    def test_to_html_format_names
      b = Workbook::Book.new([['a','b'],[1,2],[3,4]])
      c = b[0][0][0][0]
      c.format.name="testname"
      html = b.to_html
      match = html.match(/<td class=\"testname\">a<\/td>/) ? true : false
      assert_equal(true, match)
    end
    def test_to_html_css
      b = Workbook::Book.new([['a','b'],[1,2],[3,4]])
      c = b[0][0][0][0]
      c.format[:background]="#f00"
      html = b.to_html
      match = html.match(/<td>a<\/td>/) ? true : false
      assert_equal(true, match)
      html = b.to_html({:style_with_inline_css=>true})
      match = html.match(/<td style="background: #f00">a<\/td>/) ? true : false
      assert_equal(true, match)
    end
  end
end