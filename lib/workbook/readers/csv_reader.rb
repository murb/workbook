# frozen_string_literal: true

module Workbook
  module Readers
    module CsvReader
      def load_csv text, options = {}
        csv = text
        parse_csv csv, options
      end

      def csv_lib
        CSV
      end

      def parse_csv csv_raw, options = {}
        optimistic_date_converter = Workbook::Cell.new.string_optimistic_date_converter
        options = {
          converters: [optimistic_date_converter, :all],
        }.merge(options)

        csv = nil

        begin
          csv = CSV.parse(csv_raw, **options)
        rescue CSV::MalformedCSVError
          csv_excel = CSV.parse(csv_raw, **options.merge({col_sep: ";"}))
          csv = csv_excel if csv_excel[0].count > 1
        end

        if csv.nil? || (csv[0].count == 1)
          csv_excel = CSV.parse(csv_raw, **options.merge({col_sep: ";"}))
          csv = csv_excel if csv_excel[0].count > 1
        end

        self[0] = Workbook::Sheet.new(csv, self) unless sheet.has_contents?
      end
    end
  end
end
