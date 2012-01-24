module Workbook
  class Cell
    attr_accessor :value
    attr_accessor :format
    
    def initialize value=nil, options={}
      format = options[:format]
      @value = value
    end
    
    def format= f
      if f.class.is_a? Workbook::Format
        @format = f
      elsif f.class.is_a? Hash
        @format = Workbook::Format.new(f)
      elsif f.class == NilClass
        @format = nil
      end
    end
    
    def ==(other)
      if other.class==Cell
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
