# -*- encoding : utf-8 -*-
require File.join(File.dirname(__FILE__), 'helper')
module Readers
  class TestTxtReader < Test::Unit::TestCase
    # Should one day throw an error..
    # def test_failure_excel_as_txt_open
    #   w = Workbook::Book.new
    #   w.open("test/artifacts/xls_with_txt_extension.txt")
    #   puts w.sheet.table
    # end

    def test_excel_standardized_open
      w = Workbook::Book.new
      w.open("test/artifacts/excel_different_types.txt")
      # reads
      #   a,b,c,d
      # 2012-02-22,2014-12-27,2012-11-23,2012-11-12T04:20:00+00:00
      # c,222.0,,0027-12-14T05:21:00+00:00
      # 2012-01-22T11:00:00+00:00,42000.0,"goh, idd",ls
      
      assert_equal([:a,:b,:c, :d],w.sheet.table.header.to_symbols)
      assert_equal(Date.new(2012,2,22),w.sheet.table[1][:a].value)
      assert_equal("c",w.sheet.table[2][:a].value)
      assert_equal(DateTime.new(2012,1,22,11),w.sheet.table[3][:a].value)
      assert_equal(42000,w.sheet.table[3][:b].value)
      assert_equal(nil,w.sheet.table[2][:c].value)
    end
    
    def test_excel_class_read_string
      s = File.read("test/artifacts/excel_different_types.txt")
      w = Workbook::Book.read(s, :txt)
      # reads
      #   a,b,c,d
      # 2012-02-22,2014-12-27,2012-11-23,2012-11-12T04:20:00+00:00
      # c,222.0,,0027-12-14T05:21:00+00:00
      # 2012-01-22T11:00:00+00:00,42000.0,"goh, idd",ls
      
      assert_equal([:a,:b,:c, :d],w.sheet.table.header.to_symbols)
      assert_equal(Date.new(2012,2,22),w.sheet.table[1][:a].value)
      assert_equal("c",w.sheet.table[2][:a].value)
      assert_equal(DateTime.new(2012,1,22,11),w.sheet.table[3][:a].value)
      assert_equal(42000,w.sheet.table[3][:b].value)
      assert_equal(nil,w.sheet.table[2][:c].value)
    end
    
    def test_excel_instance_read_string
      s = File.read("test/artifacts/excel_different_types.txt")
      w = Workbook::Book.new
      w.read(s, :txt)
      # reads
      #   a,b,c,d
      # 2012-02-22,2014-12-27,2012-11-23,2012-11-12T04:20:00+00:00
      # c,222.0,,0027-12-14T05:21:00+00:00
      # 2012-01-22T11:00:00+00:00,42000.0,"goh, idd",ls
      
      assert_equal([:a,:b,:c, :d],w.sheet.table.header.to_symbols)
      assert_equal(Date.new(2012,2,22),w.sheet.table[1][:a].value)
      assert_equal("c",w.sheet.table[2][:a].value)
      assert_equal(DateTime.new(2012,1,22,11),w.sheet.table[3][:a].value)
      assert_equal(42000,w.sheet.table[3][:b].value)
      assert_equal(nil,w.sheet.table[2][:c].value)
    end
    
    def test_excel_instance_read_stringio
      sio = StringIO.new(File.read("test/artifacts/excel_different_types.txt"))
      w = Workbook::Book.new
      w.read(sio, :txt)
      # reads
      #   a,b,c,d
      # 2012-02-22,2014-12-27,2012-11-23,2012-11-12T04:20:00+00:00
      # c,222.0,,0027-12-14T05:21:00+00:00
      # 2012-01-22T11:00:00+00:00,42000.0,"goh, idd",ls
      
      assert_equal([:a,:b,:c, :d],w.sheet.table.header.to_symbols)
      assert_equal(Date.new(2012,2,22),w.sheet.table[1][:a].value)
      assert_equal("c",w.sheet.table[2][:a].value)
      assert_equal(DateTime.new(2012,1,22,11),w.sheet.table[3][:a].value)
      assert_equal(42000,w.sheet.table[3][:b].value)
      assert_equal(nil,w.sheet.table[2][:c].value)
    end
    
  end
end
