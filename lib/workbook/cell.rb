# encoding: utf-8
module Workbook
  class Cell
    attr_accessor :value
    attr_accessor :format
    attr_accessor :formula
    
    # Note that these types are sorted by 'importance'
    VALID_TYPES = [Numeric,String,Date,Time,TrueClass,FalseClass,NilClass]
    
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
      self_value = importance_of_class self.value
      other_value = importance_of_class other.value
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
    
  end
end
