module Workbook
  module Modules
    module TypeParser

      
      def strip_win_chars csv_raw
        csv_raw.gsub(/(\n\r|\r\n|\r)/,"\n")
      end
      
      def parse_type value
        if value.is_a? Integer
          return value
        elsif value.is_a? DateTime
          return value
        elsif value.is_a? Date
          return value
        elsif value.is_a? NilClass
          return value
        elsif value == "TRUE" or value == "true"
          return true
        elsif value == "FALSE" or value == "false"
          return false
        elsif value.to_i.to_s == value
          return value.to_i
        elsif value == ""
          return nil
        else
          return custom_date_converter.call(value)
        end
      end
      
      def custom_date_converter
        proc do |v| 
          
           begin
             if v.chars.first.to_i.to_s == v.chars.first #it should at least start with a number...
               v = (v.length > 10) ? DateTime.parse(v) : Date.parse(v) 
             end
           rescue ArgumentError
             v = v
           end
           v
         end
      end
    end
  end
end
