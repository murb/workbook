# frozen_string_literal: true
# frozen_string_literal: true

require "csv"

module Workbook
  module Writers
    module CsvTableWriter
      # Output the current workbook to CSV format
      #
      # @param [Hash] options (not used)
      # @return [String] csv (comma separated values in a string)
      def to_csv options = {}
        csv = ""
        each_with_index do |r, ri|
          line = nil
          begin
            line = CSV.generate_line(r.collect { |c| c&.value }, row_sep: "")
          rescue TypeError
            line = CSV.generate_line(r.collect { |c| c&.value })
          end
          csv += "#{line}\n"
        end
        csv
      end

      # Write the current workbook to CSV format
      #
      # @param [String] filename
      # @param [Hash] options   see #to_csv
      # @return [String] filename
      def write_to_csv filename = "untitled document.csv", options = {}
        File.open(filename, "w") { |f| f.write(to_csv(options)) }
        filename
      end
    end
  end
end
