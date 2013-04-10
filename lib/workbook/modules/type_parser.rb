# -*- encoding : utf-8 -*-
module Workbook
  module Modules
    # Adds type parsing capabilities to e.g. a Cell.
    module TypeParser
      
      # Cleans a text file from all kinds of different ways of representing new lines
      # @param [String] csv_raw a raw csv string
      def strip_win_chars csv_raw
        csv_raw.gsub(/(\n\r|\r\n|\r)/,"\n")
      end
      
      # Return the different active string parsers
      # @return [Array<Symbol>] A list of parsers
      def string_parsers
        @string_parsers ||= [:string_cleaner,:string_nil_converter,:string_integer_converter,:string_boolean_converter]
      end
      
      # Set the list of string parsers
      # @param [Array<Symbol>] parsers A list of parsers
      # @return [Array<Symbol>] A list of parsers
      def string_parsers= parsers
        @string_parsers = parsers
      end

      # Return the different active string parsers
      # @return [Array<Proc>] A list of parsers as Procs   
      def string_parsers_as_procs
        string_parsers.collect{|c| c.is_a?(Proc) ? c : self.send(c)}
      end
      
      # Returns the parsed value (retrieved by calling #value)
      # @return [Object] The parsed object, ideally a date or integer when found to be a such...
      def parse options={}
        options = {:detect_date=>false}.merge(options)
        string_parsers.push :string_optimistic_date_converter if options[:detect_date]
        v = value
        string_parsers_as_procs.each do |p|
          if v.is_a? String
            v = p.call(v)
          end
        end
        v  
      end
      
      def parse! options={}
        self.value = parse(options)
      end
      
      def clean! options={}
        parse! options
      end
      
      def string_cleaner
        proc do |v|
          v = v.strip 
          v.gsub('mailto:','')
        end
      end
      
      def string_nil_converter
        proc do |v|
          (v == "" ? nil : v)
        end
      end
      
      def string_integer_converter
        proc do |v|
          if v.to_i.to_s == v
            v.to_i
          else
            v
          end
        end
      end
      
      def string_optimistic_date_converter
        proc do |v|  
          rv = v
          starts_with_nr = v.chars.first.to_i.to_s == v.chars.first #it should at least start with a number...
          no_spaced_dash = v.to_s.match(" - ") ? false : true
          min_two_dashes = v.to_s.scan("-").count > 1 ? true : false
          min_two_dashes = v.to_s.scan("/").count > 1 ? true : false if min_two_dashes == false
          
          normal_date_length = v.to_s.length <= 25
          if no_spaced_dash and starts_with_nr and normal_date_length and min_two_dashes
            begin
              rv = (v.length > 10) ? DateTime.parse(v) : Date.parse(v) 
            rescue ArgumentError
              rv = v
            end
            begin
              rv = Date.parse(v.to_i.to_s) == rv ? v : rv # disqualify if it is only based on the first number
            rescue ArgumentError
            end
            # try strptime with format 'mm/dd/yyyy'
            if rv == v && /^\d{1,2}[\/-]\d{1,2}[\/-]\d{4}/ =~ v
              begin
                rv = Date.strptime(v, "%m/%d/%Y")
              rescue ArgumentError
              end
            end
          end          
          rv
        end
      end

      # converts 'true' or 'false' strings in `true` or `false` values
      # return [Proc] that returns a boolean value if it is considered as such 
      def string_boolean_converter
        proc do |v|
          dv = v.downcase
          if dv == "true"
            v = true
          elsif dv == "false"
            v = false
          end
          v
        end
      end
    end
  end
end
