module Workbook
  module Modules
    module TypeParser
      def strip_win_chars csv_raw
        csv_raw.gsub(/(\n\r|\r\n|\r)/,"\n")
      end
      
      def string_parsers
        @string_parsers ||= [string_cleaner,string_nil_converter,string_integer_converter,string_optimistic_date_converter,string_boolean_converter]
      end
      
      def string_parsers= arr
        @string_parsers = arr
      end
      
      def parse options={}
        options = {:detect_date=>false}.merge(options)
        v = value
        string_parsers.each do |p|
          if v.is_a? String
            v = p.call(v)
          end
        end
        v  
      end
      
      def parse! options={}
        value = parse(options)
      end
      
      def string_cleaner
        proc do |v|
          v = v.strip 
          v.gsub('mailto:','')
        end
      end
      
      def string_nil_converter
        proc do |v|
          return v == "" ? nil : v
        end
      end
      
      def string_integer_converter
        proc do |v|
          if v.to_i.to_s == v
            return v.to_i
          else
            v
          end
        end
      end
      
      def string_optimistic_date_converter
        proc do |v|  
          rv = v
           begin
             if v.chars.first.to_i.to_s == v.chars.first #it should at least start with a number...
               rv = (v.length > 10) ? DateTime.parse(v) : Date.parse(v) 
               rv = Date.parse(v.to_i.to_s) == rv ? v : rv # disqualify is it is only based on the first number
             end
           rescue ArgumentError
             rv = v
           end
           rv
         end
      end
      
      def string_boolean_converter
        proc do |v|
          dv = v.downcase
          if dv == "true" or dv == "j" or dv == "ja" or dv == "yes" or dv == "y"
            return true
          elsif dv == "false" or dv == "n" or dv == "nee" or dv == "no"
            return false
          end
          v
        end
      end
    end
  end
end
