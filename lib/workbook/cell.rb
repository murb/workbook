# encoding: utf-8
module Workbook
  class Cell
    attr_accessor :value
    attr_accessor :format
    attr_accessor :formula
    
    VALID_TYPES = [TrueClass,FalseClass,Date,Time,Numeric,String, NilClass]
    
    def valid_value? value
      valid_type = false 
      VALID_TYPES.each {|t| return true if value.is_a? t} 
      valid_type
    end
    
    def initialize value=nil, options={}     

      if valid_value? value
        format = options[:format]
        @value = value
      else
        raise ArgumentError, "value should be of a primitive type, e.g. a string, or an integer, not a #{value.class} (is_a? [TrueClass,FalseClass,Date,Time,Numeric,String, NilClass])"
      end
    end
    
    def format= f
      if f.is_a? Workbook::Format
        @format = f
      elsif f.is_a? Hash
        @format = Workbook::Format.new(f)
      elsif f.class == NilClass
        @format = nil
      end
    end

    def ==(other)
      if other.is_a? Cell
        other.value == self.value
      else
        other == self.value
      end
    end
    
    def to_sym
      value.to_s.downcase.gsub(' (j/?/leeg)','').gsub(/dd-mm-(.*)/,'').gsub(/\ja\/nee/,'').gsub(/\(\)/,'').gsub(/[\(\)]+/, "").strip.
              gsub(/(\.|\?|,|\=)/,'').
              gsub('$','').
              gsub(/\&/,'en').
              gsub(/\+/,'_plus_').
              gsub(/\s/, "_").
              gsub('–_','').
              gsub('-_','').
              gsub('+_','').
              gsub('/_','_').
              gsub('/','_').
              gsub('__','_').
              gsub('-','').
              #mb_chars.normalize(:kd).
              gsub(/[^\x00-\x7F]/n,'').downcase.to_sym
      
    end
  end
end
