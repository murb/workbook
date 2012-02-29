require 'faster_csv'

module Workbook
  module Readers
    module CsvReader
      def load_csv file_obj
        csv = file_obj.read
        csv = strip_win_chars(csv)
        parse_csv csv
        
      end
      
      def parse_csv csv_raw
        converters = [:float,:integer,:date,:date_time,custom_date_converter]
        csv=nil
        begin
          csv = FasterCSV.parse(csv_raw,{:converters=>converters})
        rescue
          # we're going to have another shot at it...
        end
        
        if csv==nil or csv[0].count == 1 
          csv_excel = FasterCSV.parse(csv_raw,{:converters=>converters,:col_sep=>';'})
          csv = csv_excel if csv_excel[0].count > 1
        end

        self[0]=Workbook::Sheet.new(csv,self) unless sheet.has_contents?
      end
 
    end
  end
end
