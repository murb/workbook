require 'test/helper'
module Modules
  class TestTableDiffSort < Test::Unit::TestCase
    def test_sort
      time = Time.now
      b = Workbook::Book.new [[1,2,3],[2,2,3],[true,false,true],["asdf","sdf","as"],[time,2,3],[2,2,2],[22,2,3],[1,2,233]]
      t = b.sheet.table
      assert_equal(
        [ [Workbook::Cell.new(1),Workbook::Cell.new(2),Workbook::Cell.new(3)],
          [Workbook::Cell.new(1),Workbook::Cell.new(2),Workbook::Cell.new(233)],
          [Workbook::Cell.new(2),Workbook::Cell.new(2),Workbook::Cell.new(2)],
          [Workbook::Cell.new(2),Workbook::Cell.new(2),Workbook::Cell.new(3)],
          [Workbook::Cell.new(22),Workbook::Cell.new(2),Workbook::Cell.new(3)],
          [Workbook::Cell.new("asdf"),Workbook::Cell.new("sdf"),Workbook::Cell.new("as")],
          [Workbook::Cell.new(time),Workbook::Cell.new(2),Workbook::Cell.new(3)],
          [Workbook::Cell.new(true),Workbook::Cell.new(false),Workbook::Cell.new(true)]
        ],
          t.sort)
      
      ba = Workbook::Book.new [['a','b','c','d'],[1,2,3,4],[4,2,3,4],[3,2,3,4]]
      sba = ba.sheet.table.sort
      assert_not_equal(Workbook::Table.new([['a','b','c','d'],[1,2,3,4],[4,2,3,4],[3,2,3,4]]),sba)
      assert_equal(Workbook::Table.new([['a','b','c','d'],[1,2,3,4],[3,2,3,4],[4,2,3,4]]),sba)
      
    end
    
    def test_align
      ba = Workbook::Book.new [['a','b','c','d'],[1,2,3,4],[4,2,3,4],[3,2,3,4]]
      bb = Workbook::Book.new [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[5,2,3,4]]
      tself = ba.sheet.table
      tother = bb.sheet.table

      placeholder_row = tother.placeholder_row

      ba = Workbook::Book.new [['a','b','c','d'],[1,2,3,4],[4,2,3,4],[3,2,3,4]]
      bb = Workbook::Book.new [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[5,2,3,4]]
      tself = ba.sheet.table
      tother = bb.sheet.table
      align_result = tself.align tother
      assert_equal("a,b,c,d\n1,2,3,4\n3,2,3,4\n4,2,3,4\n\n",align_result[:self].to_csv)  
      ba = Workbook::Book.new [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[5,2,3,4]]
      bb = Workbook::Book.new [['a','b','c','d'],[1,2,3,4],[2,2,3,4],[5,2,3,4]]
      tself = ba.sheet.table
      tother = bb.sheet.table
      align_result = tself.align tother
      assert_not_equal("a,b,c,d\n1,2,3,4\n3,2,3,4\n5,2,3,4\n\n\n",align_result[:self].to_csv)
      assert_not_equal("a,b,c,d\n1,2,3,4\n\n\n2,2,3,4\n5,2,3,4\n",align_result[:other].to_csv)
      assert_equal("a,b,c,d\n1,2,3,4\n\n3,2,3,4\n5,2,3,4\n",align_result[:self].to_csv)
      assert_equal("a,b,c,d\n1,2,3,4\n2,2,3,4\n\n5,2,3,4\n",align_result[:other].to_csv)
      
      
      
      # FOUT: 
      # "a,b,c,d\n1,2,3,4\n3,2,3,4\n5,2,3,4\n , , , \n , , , \n"
      # "a,b,c,d\n1,2,3,4\n , , , \n , , , \n2,2,3,4\n5,2,3,4\n"
      # 
      # "a,b,c,d\n1,2,3,4\n , , , \n3,2,3,4\n5,2,3,4"
      # "a,b,c,d\n1,2,3,4\n2,2,3,4\n , , , \n5,2,3,4"
      
#      assert_equal(tself.sort.insert(3,placeholder_row),align_result[:self])  
      
      
    end
    
    # def test_sort_by
    #   b = Workbook::Book.new [['a','b','c','d'],[1,2,3,4],[4,2,3,3],[3,2,3,2]]
    #   y b.sheet.table.sort_by{|r| r[:d]}
    # end
    
    def test_diff
      ba = Workbook::Book.new [['a','b','c','d'],[1,2,3,4],[4,2,3,4],[3,2,3,4],[3,3,3,4]]
      bb = Workbook::Book.new [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[5,2,3,4],[3,2,3,4]]
      # As it starts out with sorting, it is basically a comparison between            
      #  ba = [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[3,3,3,4],[4,2,3,4]]
      #  bb = [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[3,2,3,4],[5,2,3,4]]
      # then it aligns:
      #  ba = [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[3,3,3,4],[4,2,3,4],[]]
      #  bb = [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[3,2,3,4],[],[5,2,3,4]]
      # hence,
      expected = [['a','b','c','d'],[1,2,3,4],[3,2,3,4],[3,'3 (was: 2)',3,4],[4,2,3,4],['(was: 5)','(was: 2)','(was: 3)','(was: 4)']]
      
      tself = ba.sheet.table
      tother = bb.sheet.table
      diff_result = tself.diff tother
      assert_equal('a',diff_result.sheet.table.header[0].value)
      assert_equal({:rotation=>72, :font_weight=>:bold},diff_result.sheet.table.header[0].format)
      assert_equal(expected[2][0],diff_result.sheet.table[2][0].value)
      assert_equal(expected[3][1],diff_result.sheet.table[3][1].value) # 3 (was: 2)
      assert_equal(expected[4][2],diff_result.sheet.table[4][2].value) # 4
      assert_equal(expected[5][3],diff_result.sheet.table[5][3].value) # (was: 3)
      diff_result.write_to_xls({:rewrite_header=>true})
    end
    
    
    
    def test_diff_xls
      prev = "test/artifacts/compare_old.xls"
      curr = "test/artifacts/compare_current.xls"
      
      wprev=Workbook::Book.new
      wprev.load_xls prev
      wcurr=Workbook::Book.new
      wcurr.load_xls curr
      puts "\nStart diff"
      diff = wcurr.sheet.table.diff wprev.sheet.table
      puts "Start writing"
      puts diff.sheet.table.to_csv
    end
  end
end
