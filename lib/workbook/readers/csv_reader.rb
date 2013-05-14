# -*- encoding : utf-8 -*-

module Workbook
  module Readers
    module CsvReader
      def load_csv text
        csv = text
        parse_csv csv
      end

      def csv_lib
        if RUBY_VERSION < '1.9'
          require 'faster_csv'
          return FasterCSV
        else
          return CSV
        end
      end

      def parse_csv csv_raw
        custom_date_converter = Workbook::Cell.new.string_optimistic_date_converter
        converters = [:float,:integer,:date,:date_time,custom_date_converter]
        csv=nil
        #begin
        csv = csv_lib.parse(csv_raw,{:converters=>converters})

        #rescue
        # we're going to have another shot at it...
        #end

        if csv==nil or csv[0].count == 1
          csv_excel = csv_lib.parse(csv_raw,{:converters=>converters,:col_sep=>';'})
          csv = csv_excel if csv_excel[0].count > 1
        end

        self[0]=Workbook::Sheet.new(csv,self) unless sheet.has_contents?
      end

    end
  end
end
