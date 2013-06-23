# -*- encoding : utf-8 -*-
require 'csv'

module Workbook
  module Writers
    module CsvTableWriter
      # Output the current workbook to CSV format
      #
      # @param [String] filename
      # @param [Hash] options (not used)
      # @return [String] csv (comma separated values in a string)
      def to_csv options={}
        csv = ""
        options = {}.merge options
        self.each_with_index do |r, ri|
          line=nil
          begin
            line = CSV::generate_line(r.collect{|c| c.value if c},{:row_sep=>""})
          rescue TypeError
            line = CSV::generate_line(r.collect{|c| c.value if c})
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
      def write_to_csv filename="#{title}.csv", options={}
        File.open(filename, 'w') {|f| f.write(to_csv(options)) }
        return filename
      end

    end
  end
end
