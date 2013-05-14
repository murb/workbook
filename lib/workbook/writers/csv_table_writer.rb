# -*- encoding : utf-8 -*-
require 'csv'

module Workbook
  module Writers
    module CsvTableWriter
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

    end
  end
end
