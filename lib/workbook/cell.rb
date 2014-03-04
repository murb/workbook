# -*- encoding : utf-8 -*-

require 'workbook/modules/type_parser'

module Workbook
  class Cell
    include Workbook::Modules::TypeParser

    attr_accessor :formula

    # Note that these types are sorted by 'importance'
    VALID_TYPES = [Numeric,String,Time,Date,TrueClass,FalseClass,NilClass,Workbook::NilValue]

    # Evaluates a value for class-validity
    #
    # @param [Numeric,String,Time,Date,TrueClass,FalseClass,NilClass,Object] value the value to evaluate
    # @return [Boolean] returns true when the value is a valid cell value
    def valid_value? value
      valid_type = false
      VALID_TYPES.each {|t| return true if value.is_a? t}
      valid_type
    end

    # @param [Numeric,String,Time,Date,TrueClass,FalseClass,NilClass] value a valid value
    # @param [Hash] options a reference to :format (Workbook::Format) can be specified
    def initialize value=nil, options={}
      if valid_value? value
        self.format = options[:format]
        @value = value
        @to_sym = nil
      else
        raise ArgumentError, "value should be of a primitive type, e.g. a string, or an integer, not a #{value.class} (is_a? [TrueClass,FalseClass,Date,Time,Numeric,String, NilClass])"
      end
    end

    # Change the current value
    #
    # @param [Numeric,String,Time,Date,TrueClass,FalseClass,NilClass] value a valid value
    def value= value
      if valid_value? value
        @value = value
        @to_sym = nil
      else
        raise ArgumentError, "value should be of a primitive type, e.g. a string, or an integer, not a #{value.class} (is_a? [TrueClass,FalseClass,Date,Time,Numeric,String, NilClass])"
      end
    end

    # Returns the current value
    #
    # @return [Numeric,String,Time,Date,TrueClass,FalseClass,NilClass] a valid value
    def value
      @value
    end

    # Change the current format
    #
    # @param [Workbook::Format, Hash] f set the formatting properties of this Cell
    def format= f
      if f.is_a? Workbook::Format
        @format = f
      elsif f.is_a? Hash
        @format = Workbook::Format.new(f)
      elsif f.class == NilClass
        @format = Workbook::Format.new
      end
    end

    # Returns current format
    #
    # @returns [Workbook::Format] the current format
    def format
      @format ||= Workbook::Format.new
    end

    # Tests for equality based on its value (formatting is irrelevant)
    #
    # @param [Workbook::Cell] other cell to compare against
    # @return [Boolean]
    def ==(other)
      if other.is_a? Cell
        other.value == self.value
      else
        other == self.value
      end
    end

    # returns true when the value of the cell is nil.
    # @return [Boolean]
    def nil?
      return value.nil?
    end

    # returns a symbol representation of the cell's value
    # @return [Symbol] a symbol representation
    # @example
    #
    #     <Workbook::Cell value="yet another value">.to_sym # returns :yet_another_value
    def to_sym
      return @to_sym if @to_sym
      #mb_chars.normalize(:kd).
      v = nil
      if value
        ends_with_exclamationmark = (value[-1] == '!')
        ends_with_questionmark = (value[-1] == '?')
        v = value.to_s.downcase
        v = v.strip.gsub(/[\(\)\.\?\,\!\=\$]/,'').
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

        accents = {
          ['á','à','â','ä','ã','å'] => 'a',
          ['Ã','Ä','Â','À','�?','Å'] => 'A',
          ['é','è','ê','ë'] => 'e',
          ['Ë','É','È','Ê'] => 'E',
          ['í','ì','î','ï'] => 'i',
          ['�?','Î','Ì','�?'] => 'I',
          ['ó','ò','ô','ö','õ'] => 'o',
          ['Õ','Ö','Ô','Ò','Ó'] => 'O',
          ['ú','ù','û','ü'] => 'u',
          ['Ú','Û','Ù','Ü'] => 'U',
          ['ç'] => 'c',
          ['Ç'] => 'C',
          ['š', 'ś'] => 's',
          ['Š', 'Ś'] => 'S',
          ['ž','ź'] => 'z',
          ['Ž','Ź'] => 'Z',
          ['ñ'] => 'n',
          ['Ñ'] => 'N'
        }
        accents.each do |ac,rep|
          ac.each do |s|
            v = v.gsub(s, rep)
          end
        end
        if RUBY_VERSION < '1.9'
          v = v.gsub(/[^\x00-\x7F]/n,'')
        else
          # See String#encode
          encoding_options = {:invalid => :replace, :undef => :replace, :replace => ''}
          v = v.encode(Encoding.find('ASCII'), encoding_options)
        end
        v = "#{v}!" if ends_with_exclamationmark
        v = "#{v}?" if ends_with_questionmark
        v = v.downcase.to_sym
      end
      @to_sym = v
      return @to_sym
    end

    # Compare
    # @param [Workbook::Cell] other cell to compare against (based on value), can compare different value-types using #compare_on_class
    # @returns [Fixnum] -1, 0, 1
    def <=> other
      rv = nil
      begin
        rv = self.value <=> other.value
      rescue NoMethodError
        rv = compare_on_class other
      end
      if rv == nil
        rv = compare_on_class other
      end
      return rv

    end

    # Compare on class level
    #
    # @param [Workbook::Cell] other cell to compare against
    def compare_on_class other
      other_value = nil
      other_value = other.value if other
      self_value = importance_of_class self.value
      other_value = importance_of_class other_value
      self_value <=> other_value
    end

    # Returns the importance of a value's class
    #
    # @param value a potential value for a cell
    def importance_of_class value
      VALID_TYPES.each_with_index do |c,i|
        return i if value.is_a? c
      end
      return nil
    end

    def inspect
      "<Workbook::Cell @value=#{value}>"
    end

    # convert value to string, and in case of a Date or Time value, apply formatting
    # @return [String]
    def to_s
      if (value.is_a? Date or value.is_a? Time) and format[:number_format]
        value.strftime(format[:number_format])
      else
        value.to_s
      end
    end

    def colspan= c
      @colspan = c
    end
    def rowspan= r
      @rowspan = r
    end

    def colspan
      @colspan.to_i if defined?(@colspan) and @colspan.to_i > 1
    end
    def rowspan
      @rowspan.to_i if defined?(@rowspan) and @rowspan.to_i > 1
    end
  end
end
