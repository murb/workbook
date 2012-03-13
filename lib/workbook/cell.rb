# encoding: utf-8

require 'workbook/modules/type_parser'

module Workbook
  class Cell
    include Workbook::Modules::TypeParser
    
    attr_accessor :value
    attr_accessor :format
    attr_accessor :formula
    
    # Note that these types are sorted by 'importance'
    VALID_TYPES = [Numeric,String,Time,Date,TrueClass,FalseClass,NilClass]
    
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
        @format = Workbook::Format.new
      end
    end
    
    def format
      @format ||= Workbook::Format.new
    end

    def ==(other)
      if other.is_a? Cell
        other.value == self.value
      else
        other == self.value
      end
    end
    
    def nil?
      return value.nil?
    end
    
    def to_sym
      #mb_chars.normalize(:kd).
      v = nil
      if value
        v = value.to_s.downcase
        v = v.gsub(' (j/?/leeg)','').gsub(/dd-mm-(.*)/,'').gsub(/\ja\/nee/,'').gsub(/\(\)/,'').gsub(/[\(\)]+/, '')
        v = v.strip.gsub(/(\.|\?|,|\=)/,'').
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
            gsub('-','')
        v = v.gsub(/[^\x00-\x7F]/n,'').downcase.to_sym
      end
      v
    end
    
    def <=> other
      rv = nil
      begin
        rv = self.value <=> other.value
      rescue NoMethodError => e
        rv = compare_on_class other
      end
      if rv == nil        
        rv = compare_on_class other
      end
      return rv

    end
    
    def compare_on_class other
      other_value = nil
      other_value = other.value if other
      self_value = importance_of_class self.value
      other_value = importance_of_class other_value
      self_value <=> other_value
    end
    
    def importance_of_class value
      VALID_TYPES.each_with_index do |c,i|  
        return i if value.is_a? c
      end
      return nil
    end
    
    def inspect
      "<Workbook::Cell @value=#{value}>"
    end
    
    def to_s
      if (value.is_a? Date or value.is_a? Time) and format[:number_format]
        value.strftime(format[:number_format])
      else
        value.to_s
      end
    end
    
  end
end
