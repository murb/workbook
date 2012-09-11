require File.join(File.dirname(__FILE__), 'helper')
module Readers
  class TestTxtReader < Test::Unit::TestCase
    # Should one day throw an error..
    # def test_failure_excel_as_txt_open
    #   w = Workbook::Book.new
    #   w.open("test/artifacts/xls_with_txt_extension.txt")
    #   puts w.sheet.table
    # end
    def test_complex_excel_txt_open
      w = Workbook::Book.new
      w.open("test/artifacts/sharepoint_download_excel.txt")
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
      assert_equal(Date.new(2011,11,01),w.sheet.table[29][:gewenste_migratie_maand].value)
      assert_equal("Belt , G (Gerrit)",w.sheet.table[1][:gemaakt_door].value)
      assert_equal(false,w.sheet.table[1][:geen_aanbod_voor_isp].value)
      assert_equal(DateTime.new(2011,7,25,11,02),w.sheet.table[4][:gemaakt].value)
      assert_equal("Roggeveen , PC (Peter)",w.sheet.table[4][:gemaakt_door].value)
      assert_equal("Groeneveld, R (René)",w.sheet.table[400][:gemaakt_door].value)
      assert_equal(24208,w.sheet.table[7][:locatie].value)
    end
    
  end
end
