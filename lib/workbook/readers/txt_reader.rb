$KCODE="u"
require 'faster_csv'

module Workbook
  module Readers
    module TxtReader
      def load_txt file_obj
        csv = file_obj.read
        # TODO: while this will do for now, should be a more comprehensive replace: http://sh.codefetch.com/example/h9/sh/iso8859-1-to-html.sh
        # TODO: not sure whether this is ok for 1.9.x compatibility; focus on parsing windows excel sheets
        csv = csv.gsub(/\351/,"é").gsub(/\240/," ").gsub(/\353/,"ë").gsub(/\357/,"ï").gsub(/\377/," ")
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