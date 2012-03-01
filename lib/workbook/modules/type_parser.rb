module Workbook
  module Modules
    module TypeParser

      
      def strip_win_chars csv_raw
        csv_raw.gsub(/(\n\r|\r\n|\r)/,"\n")
      end
      
      def parse_type value, options={}
        options = {:detect_date=>true}
        if value.is_a? Integer
          return value
        elsif value.is_a? DateTime
          return value
        elsif value.is_a? Date
          return value
        elsif value.is_a? NilClass
          return value
        elsif value.to_i.to_s == value
          return value.to_i
        elsif value.is_a? String
          value = value.strip 
          value.gsub('mailto:','')
          value = string_to_boolean value
          value
        end
        if value == ""
          return nil
        elsif options[:detect_date] == true
          return custom_date_converter.call(value)
        end
        value
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
      
      def string_to_boolean sss
        dvalue = sss.downcase
        if dvalue == "true" or dvalue == "j" or dvalue == "ja" or dvalue == "yes" or dvalue == "y"
          return true
        elsif dvalue == "false" or dvalue == "n" or dvalue == "nee" or dvalue == "no"
          return false
        end
        sss
      end
    end
  end
end
