# -*- encoding : utf-8 -*-

module Workbook
  module Readers
    module CsvReader
      def load_csv text, options={}
        csv = text
        parse_csv csv, options
      end

      def csv_lib
        if RUBY_VERSION < '1.9'
          require 'faster_csv'
          return FasterCSV
        else
          return CSV
        end
      end

      def parse_csv csv_raw, options={}
        custom_date_converter = Workbook::Cell.new.string_optimistic_date_converter
        options = {
          converters: [:float,:integer,:date,:date_time,custom_date_converter]
        }.merge(options)

        csv=nil
        begin
          csv = csv_lib.parse(csv_raw,options)

        rescue CSV::MalformedCSVError
          csv_excel = csv_lib.parse(csv_raw,options.merge({:col_sep=>';'}))
          csv = csv_excel if csv_excel[0].count > 1

        end

        if csv==nil or csv[0].count == 1
          csv_excel = csv_lib.parse(csv_raw,options.merge({:col_sep=>';'}))
          csv = csv_excel if csv_excel[0].count > 1
        end

        self[0]=Workbook::Sheet.new(csv,self) unless sheet.has_contents?
      end

    end
  end
end
