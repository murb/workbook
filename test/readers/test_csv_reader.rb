require File.join(File.dirname(__FILE__), '../helper')
module Readers
  class TestCsvWriter < Test::Unit::TestCase
    def test_open
      w = Workbook::Book.new
      w.open 'test/artifacts/simple_csv.csv'
      # reads
      #       a,b,c,d
      #       1,2,3,4
      #       5,3,2,1
      #       "asdf",123,12,2001-02-02
      #       
      assert_equal([:a,:b,:c,:d],w.sheet.table.header.to_symbols)
      assert_equal(3,w.sheet.table[2][:b].value)
      assert_equal("asdf",w.sheet.table[3][:a].value)
      assert_equal(Date.new(2001,2,2),w.sheet.table[3][:d].value)
      
    end
    def test_excel_csv_open
      w = Workbook::Book.new
      w.open("test/artifacts/simple_excel_csv.csv")
      # reads
      #   a;b;c
      #   1-1-2001;23;1
      #   asdf;23;asd
      #   23;asdf;sadf
      #   12;23;12-02-2011 12:23
      #y w.sheet.table
      assert_equal([:a,:b,:c],w.sheet.table.header.to_symbols)
      assert_equal(23,w.sheet.table[2][:b].value)
      assert_equal("sadf",w.sheet.table[3][:c].value)
      assert_equal(Date.new(2001,1,1),w.sheet.table[1][:a].value)
      assert_equal(DateTime.new(2011,2,12,12,23),w.sheet.table[4][:c].value)
    end
    
    def test_complex_excel_csv_open
      w = Workbook::Book.new
      w.open("test/artifacts/sharepoint_download_excel.csv")
      # reads
      #   a;b;c
      #   1-1-2001;23;1
      #   asdf;23;asd
      #   23;asdf;sadf
      #   12;23;12-02-2011 12:23
      #y w.sheet.table
      assert_equal([:id,
 :naam,
 :locatie_naam,
 :locatie_rayon,
 :datum_gesprek,
 :algehele_indruk,
 :gemaakt_door,
 :gemaakt,
 :titel,
 :toelichting,
 :gewenste_migratie_maand,
 :plek_ing_balie,
 :plek_tnt_balie,
 :alternatieve_locaties,
 :datum_introductie_gesprek,
 :geen_aanbod_voor_isp,
 :oplevering_geslaagd_op,
 :nazorgpunten,
 :locatie],w.sheet.table.header.to_symbols)
 #puts w.sheet.table.to_csv
      assert_equal(Date.new(2012,7,27),w.sheet.table[2][:gewenste_migratie_maand].value)
      assert_equal("sadf",w.sheet.table[22][:datum_introductie_gesprek].value.to_s)
      assert_equal(Date.new(2001,1,1),w.sheet.table[1][:geen_aanbod_voor_isp].value.to_s)
      assert_equal(DateTime.new(2011,2,12,12,23),w.sheet.table[4][:gemaakt].value.to_s)
      assert_equal(DateTime.new(2011,2,12,12,23),w.sheet.table[4][:gemaakt_door].value.to_s)
    end
    
  end
end
