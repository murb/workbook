require 'faster_csv'

module Workbook
  module Readers
    module XlsReader
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
        begin
          if csv==nil or csv[0].count == 1 
            csv_excel = FasterCSV.parse(csv_raw,{:converters=>converters,:col_sep=>';'})
            csv = csv_excel if csv_excel[0].count > 1
          end
        rescue
          puts "Warning: Slower CSV based parsing"
          
          csv = []
          csv_raw.each do |line|
            # While slower, CSV is a bit more forgiving
            parsed = CSV.parse_line(line.to_s.gsub(';',','))
            parsed.convert
          end
        end
        self[0]=Workbook::Sheet.new(csv,self) unless sheet.has_contents?
      end
      
      def strip_win_chars csv_raw
        csv_raw.gsub(/(\n\r|\r\n|\r)/,"\n")
      end
      
      def custom_date_converter
        proc do |v| 
           begin
             v = (v.length > 10) ? DateTime.parse(v) : Date.parse(v) 
           rescue ArgumentError
             v = v
           end
           v
         end
      end
    end
  end
end
