require 'CSV'

module Workbook
  module Writers
    module CsvTableWriter
      def to_csv options={}
        csv = ""
        options = {}.merge options
        self.each_with_index do |r, ri|
          line = CSV::generate_line(r.collect{|c| c.value if c})
          csv += "#{line}\n"
        end
        csv
      end
      
    end
  end
end