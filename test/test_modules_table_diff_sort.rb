# frozen_string_literal: true

require File.join(File.dirname(__FILE__), "helper")
module Modules
  class TestTableDiffSort < Minitest::Test
    def test_sort
      time = Time.now
      b = Workbook::Book.new [[1, 2, 3], [2, 2, 3], [true, false, true], ["asdf", "sdf", "as"], [time, 2, 3], [2, 2, 2], [22, 2, 3], [1, 2, 233]]
      t = b.sheet.table
      assert_equal(
        [[Workbook::Cell.new(1), Workbook::Cell.new(2), Workbook::Cell.new(3)],
         [Workbook::Cell.new(1), Workbook::Cell.new(2), Workbook::Cell.new(233)],
         [Workbook::Cell.new(2), Workbook::Cell.new(2), Workbook::Cell.new(2)],
         [Workbook::Cell.new(2), Workbook::Cell.new(2), Workbook::Cell.new(3)],
         [Workbook::Cell.new(22), Workbook::Cell.new(2), Workbook::Cell.new(3)],
         [Workbook::Cell.new("asdf"), Workbook::Cell.new("sdf"), Workbook::Cell.new("as")],
         [Workbook::Cell.new(time), Workbook::Cell.new(2), Workbook::Cell.new(3)],
         [Workbook::Cell.new(true), Workbook::Cell.new(false), Workbook::Cell.new(true)],],
      t.sort.rows
      )

      ba = Workbook::Book.new [["a", "b", "c", "d"], [1, 2, 3, 4], [4, 2, 3, 4], [3, 2, 3, 4]]
      sba = ba.sheet.table.sort
      refute_equal(Workbook::Table.new([["a", "b", "c", "d"], [1, 2, 3, 4], [4, 2, 3, 4], [3, 2, 3, 4]]), sba)
      assert_equal(Workbook::Table.new([["a", "b", "c", "d"], [1, 2, 3, 4], [3, 2, 3, 4], [4, 2, 3, 4]]), sba)
    end

    def test_align
      ba = Workbook::Book.new [["a", "b", "c", "d"], [1, 2, 3, 4], [4, 2, 3, 4], [3, 2, 3, 4]]
      bb = Workbook::Book.new [["a", "b", "c", "d"], [1, 2, 3, 4], [3, 2, 3, 4], [5, 2, 3, 4]]
      tself = ba.sheet.table
      tother = bb.sheet.table
      align_result = tself.align tother
      assert_equal("a,b,c,d\n1,2,3,4\n3,2,3,4\n\n5,2,3,4\n", align_result[:other].to_csv)
      assert_equal("a,b,c,d\n1,2,3,4\n3,2,3,4\n4,2,3,4\n\n", align_result[:self].to_csv)

      ba = Workbook::Book.new [["a", "b", "c", "d"], [1, 2, 3, 4], [3, 2, 3, 4], [5, 2, 3, 4]]
      bb = Workbook::Book.new [["a", "b", "c", "d"], [1, 2, 3, 4], [2, 2, 3, 4], [5, 2, 3, 4]]
      tself = ba.sheet.table
      tother = bb.sheet.table
      align_result = tself.align tother
      assert_equal("a,b,c,d\n1,2,3,4\n\n3,2,3,4\n5,2,3,4\n", align_result[:self].to_csv)
      assert_equal("a,b,c,d\n1,2,3,4\n2,2,3,4\n\n5,2,3,4\n", align_result[:other].to_csv)
      align_result = tother.align tself
      assert_equal("a,b,c,d\n1,2,3,4\n\n3,2,3,4\n5,2,3,4\n", align_result[:other].to_csv)
      assert_equal("a,b,c,d\n1,2,3,4\n2,2,3,4\n\n5,2,3,4\n", align_result[:self].to_csv)

      tself =  Workbook::Book.new([["a", "b", "c", "d"], [1, 2, 3, 4], [1, 3, 3, 4], [3, 2, 3, 4], [5, 2, 3, 4]]).sheet.table
      tother = Workbook::Book.new([["a", "b", "c", "d"], [1, 2, 3, 4], [2, 2, 3, 4], [5, 2, 3, 4]]).sheet.table
      align_result = tself.align tother
      assert_equal("a,b,c,d\n1,2,3,4\n1,3,3,4\n\n3,2,3,4\n5,2,3,4\n", align_result[:self].to_csv)
      assert_equal("a,b,c,d\n1,2,3,4\n\n2,2,3,4\n\n5,2,3,4\n", align_result[:other].to_csv)
      tself =  Workbook::Book.new([["a", "b", "c", "d"], [1, 2, 3, 4],                                           [3, 2, 3, 4], [5, 2, 3, 4]]).sheet.table
      tother = Workbook::Book.new([["a", "b", "c", "d"], [1, 2, 3, 4], [1, 2, 3, 4], [1, 2, 3, 4], [2, 2, 3, 4],               [5, 2, 3, 4]]).sheet.table
      align_result = tself.align tother
      assert_equal("a,b,c,d\n1,2,3,4\n\n\n\n3,2,3,4\n5,2,3,4\n", align_result[:self].to_csv)
      assert_equal("a,b,c,d\n1,2,3,4\n1,2,3,4\n1,2,3,4\n2,2,3,4\n\n5,2,3,4\n", align_result[:other].to_csv)
      tself = Workbook::Book.new([["a", "b", "c", "d"], [9, 2, 3, 4], [3, 2, 3, 4], [5, 2, 3, 4]]).sheet.table
      tother = Workbook::Book.new([["a", "b", "c", "d"], [1, 2, 3, 4], [1, 2, 3, 4], [1, 2, 3, 4], [2, 2, 3, 4], [5, 2, 3, 4]]).sheet.table
      align_result = tself.align tother
      assert_equal("a,b,c,d\n\n\n\n\n3,2,3,4\n5,2,3,4\n9,2,3,4\n", align_result[:self].to_csv)
      assert_equal("a,b,c,d\n1,2,3,4\n1,2,3,4\n1,2,3,4\n2,2,3,4\n\n5,2,3,4\n\n", align_result[:other].to_csv)
      tself = Workbook::Book.new([["a", "b", "c", "d"], [49, 2, 3, 4], [10, 2, 3, 4], [45, 2, 3, 4]]).sheet.table
      tother = Workbook::Book.new([["a", "b", "c", "d"], [1, 2, 3, 4], [1, 2, 3, 4], [1, 2, 3, 4], [2, 2, 3, 4], [5, 2, 3, 4]]).sheet.table
      align_result = tself.align tother
      assert_equal("a,b,c,d\n\n\n\n\n\n10,2,3,4\n45,2,3,4\n49,2,3,4\n", align_result[:self].to_csv)
      assert_equal("a,b,c,d\n1,2,3,4\n1,2,3,4\n1,2,3,4\n2,2,3,4\n5,2,3,4\n\n\n\n", align_result[:other].to_csv)
    end

    def test_with_nil_header
      a = Workbook::Book.new([["a", "b", nil], [1, 2, 3]])
      b = Workbook::Book.new([["a", "b", "c"], [1, 2, 3]])
      c = a.sheet.table.diff b.sheet.table
      assert_equal("a,b,(was: c)\n1,2,(was: 3)\n", c.to_csv)
    end

    # def test_diff_file
    #   a = Workbook::Book.open('test/artifacts/private_a_0.xls')
    #   b = Workbook::Book.open('test/artifacts/private_a_0.xls')
    #   y a.sheet.table.diff b.sheet.table
    # end

    # def test_sort_by
    #   b = Workbook::Book.new [['a','b','c','d'],[1,2,3,4],[4,2,3,3],[3,2,3,2]]
    #   y b.sheet.table.sort_by{|r| r[:d]}
    # end

    def test_diff
      ba = Workbook::Book.new [["a", "b", "c", "d"], [1, 2, 3, 4], [4, 2, 3, 4], [3, 2, 3, 4], [3, 3, 3, 4]]
      bb = Workbook::Book.new [["a", "b", "c", "d"], [1, 2, 3, 4], [3, 2, 3, 4], [5, 2, 3, 4], [3, 2, 3, 4]]
      # Start:
      # ba = [['a','b','c','d'],[1,2,3,4],[4,2,3,4],[3,2,3,4],[3,3,3,4]]
      # bb = [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[5,2,3,4],[3,2,3,4]]
      # As it starts out with sorting, it is basically a comparison between
      #  ba = [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[3,3,3,4],[4,2,3,4],[]]
      #  bb = [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[3,2,3,4],[],[5,2,3,4]]
      # then it aligns:
      #  ba = [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[3,3,3,4],[4,2,3,4],[]]
      #  bb = [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[3,2,3,4],[],[5,2,3,4]]
      # hence,

      tself = ba.sheet.table
      tother = bb.sheet.table
      diff_result = tself.diff tother
      assert_equal("a", diff_result.sheet.table.header[0].value)
      assert_equal("a,b,c,d\n1,2,3,4\n3,2,3,4\n3,3 (was: 2),3,4\n4,2,3,4\n(was: 5),(was: 2),(was: 3),(was: 4)\n", diff_result.to_csv)
      diff_result.sheet.book.write_to_xls
    end

    def test_workbook_diff
      ba = Workbook::Book.new [["a", "b", "c", "d"], [1, 2, 3, 4], [4, 2, 3, 4], [3, 2, 3, 4], [3, 3, 3, 4]]
      bb = Workbook::Book.new [["a", "b", "c", "d"], [1, 2, 3, 4], [3, 2, 3, 4], [5, 2, 3, 4], [3, 2, 3, 4]]
      ba << [["e", "f", "g", "h"], [1, 2, 3, 4], [4, 2, 3, 4], [3, 2, 3, 4], [3, 3, 3, 4]]
      bb << [["e", "f", "g", "h"], [1, 2, 2, 2], [4, 2, 3, 4], [3, 2, 3, 4], [3, 3, 3, 6]]
      diff_result = ba.diff bb

      assert_equal("a,b,c,d\n1,2,3,4\n3,2,3,4\n3,3 (was: 2),3,4\n4,2,3,4\n(was: 5),(was: 2),(was: 3),(was: 4)\n", diff_result[0][0].to_csv)
      assert_equal("e,f,g,h\n1,2,3 (was: 2),4 (was: 2)\n3,2,3,4\n3,3,3,4 (was: 6)\n4,2,3,4\n", diff_result[1][0].to_csv)
    end
  end
end
