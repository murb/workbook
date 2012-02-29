require 'faster_csv'

module Workbook
  module Readers
    module TxtReader
      def load_txt file_obj
        csv = file_obj.read
        #while this will do for now, should be a more comprehensive replace: http://sh.codefetch.com/example/h9/sh/iso8859-1-to-html.sh
        csv = csv.gsub(/\351/,"é").gsub(/\240/," ").gsub(/\353/,"ë").gsub(/\357/,"ï")
        parse_txt csv
        
      end
      
      def parse_txt csv_raw
        csv = []
        csv_raw.split("\r\n").each {|l| csv << FasterCSV.parse_line(l,{:col_sep=>"\t"}).collect{|c| parse_type(c)} };nil
        self[0]=Workbook::Sheet.new(csv,self) unless sheet.has_contents?
      end

      # def parse value
      #    if value.is_a? Integer
      #      return value
      #    elsif value.is_a? DateTime
      #      return value
      #    elsif value.is_a? Date
      #      return value
      #    elsif value.is_a? NilClass
      #      return value
      #    elsif value == "TRUE"
      #      return true
      #    elsif value == "FALSE"
      #      return false
      #    elsif value == ""
      #      return nil
      #    else
      #      return custom_date_converter.call(value)
      #    end
      #  end
      
      # def custom_date_converter
      #        proc do |v| 
      #           begin
      #             v = (v.length > 10) ? DateTime.parse(v) : Date.parse(v) 
      #           rescue ArgumentError
      #             v = v
      #           end
      #           v
      #         end
      #      end
    end
  end
end