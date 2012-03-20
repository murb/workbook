require 'faster_csv'
module Workbook
  module Readers
    module TxtReader
      def load_txt text
        csv = text
        parse_txt csv        
      end
      
      def parse_txt csv_raw
        csv = []
        csv_raw.split("\r\n").each {|l| csv << FasterCSV.parse_line(l,{:col_sep=>"\t"});nil}
        self[0]=Workbook::Sheet.new(csv,self,{:parse_cells_on_batch_creation=>true, :cell_parse_options=>{:detect_date=>true}}) unless sheet.has_contents?
      end
    end
  end
end