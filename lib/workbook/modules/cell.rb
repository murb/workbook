# -*- encoding : utf-8 -*-
require 'workbook/modules/type_parser'
require 'workbook/nil_value'

module Workbook
  module Modules
    module Cell
      include Workbook::Modules::TypeParser



      # Note that these types are sorted by 'importance'
      VALID_TYPES = [Numeric,String,Time,Date,TrueClass,FalseClass,NilClass,Workbook::NilValue,Symbol]

      # Evaluates a value for class-validity
      #
      # @param [Numeric,String,Time,Date,TrueClass,FalseClass,NilClass,Object] value the value to evaluate
      # @return [Boolean] returns true when the value is a valid cell value
      def valid_value? value
        valid_type = false
        VALID_TYPES.each {|t| return true if value.is_a? t}
        valid_type
      end

      def formula
        @formula
      end

      def formula= f
        @formula = f
      end

      def row
        @row
      end

      def row= r
        @row= r
      end

      # @param [Numeric,String,Time,Date,TrueClass,FalseClass,NilClass] value a valid value
      # @param [Hash] options a reference to :format (Workbook::Format) can be specified
      def initialize value=nil, options={}
        self.format = options[:format] if options[:format]
        self.row = options[:row]
        self.value = value
        @to_sym = nil
      end

      # Change the current value
      #
      # @param [Numeric,String,Time,Date,TrueClass,FalseClass,NilClass,Symbol] value a valid value
      def value= value
        if valid_value? value
          @value = value
          @to_sym = nil
        else
          raise ArgumentError, "value should be of a primitive type, e.g. a string, or an integer, not a #{value.class} (is_a? [TrueClass,FalseClass,Date,Time,Numeric,String, NilClass, Symbol])"
        end
      end

      # Returns column type, either :primary_key, :string, :text, :integer, :float, :decimal, :datetime, :date, :binary, :boolean
      #
      # @return [Symbol] the type of cell, compatible with Workbook::Column'types
      def cell_type
        tp = case value.class.to_s
        when "String" then :string
        when "FalseClass" then :boolean
        when "TrueClass" then :boolean
        when 'Time' then :time
        when "Date" then :date
        when "DateTime" then :datetime
        when "Float" then :float
        when "Integer" then :integer
        when "Numeric" then :integer
        when "Fixnum" then :integer
        when "Symbol" then :string
        end
      end


      # Returns the current value
      #
      # @return [Numeric,String,Time,Date,TrueClass,FalseClass,NilClass] a valid value
      def value
        @value
      end

      # Returns the sheet its at.
      #
      # @return [Workbook::Table]
      def table
        row.table if row
      end

      # Quick assessor to the book's template, if it exists
      #
      # @return [Workbook::Template]
      def template
        row.template if row
      end

      # Change the current format
      #
      # @param [Workbook::Format, Hash] f set the formatting properties of this Cell, see Workbook::Format#initialize
      def format= f
        if f.is_a? Workbook::Format
          @workbook_format = f
        elsif f.is_a? Hash
          @workbook_format = Workbook::Format.new(f)
        elsif f.class == NilClass
          @workbook_format = Workbook::Format.new
        end
      end

      # Returns current format
      #
      # @return [Workbook::Format] the current format
      def format
        # return @workbook_format if @workbook_format
        if row and template and row.header? and !@workbook_format
          @workbook_format = template.create_or_find_format_by(:header)
        else
          @workbook_format ||= Workbook::Format.new
        end
        @workbook_format
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
        v = nil
        if value
          v = value.to_s.downcase
          if v.to_i != 0
            v = "num#{v}"
          end
          ends_with_exclamationmark = (v[-1] == '!')
          ends_with_questionmark = (v[-1] == '?')

          replacements = {
            [/[\(\)\.\?\,\!\=\$\:]/,] => '',
            [/\&/] => 'amp',
            [/\+/] => '_plus_',
            [/\s/,'/_','/',"\\"] => '_',
            ['–_','-_','+_','-'] => '',
            ['__']=>'_',
            ['>']=>'gt',
            ['<']=>'lt',
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
            ['Ñ'] => 'N',
            ['#'] => 'hash',
            ['*'] => 'asterisk'
          }
          replacements.each do |ac,rep|
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
      #
      # @param [Workbook::Cell] other cell to compare against (based on value), can compare different value-types using #compare_on_class
      # @return [Fixnum] -1, 0, 1
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

      # Returns whether special formatting is present on this cell
      #
      # @return [Boolean] index of the cell
      def format?
        format and format.keys.count > 0
      end

      # Returns the index of the cell within the row, returns nil if no row is present
      #
      # @return [Integer, NilClass] index of the cell
      def index
        row.index self if row
      end

      # Returns the key (a Symbol) of the cell, based on its table's header
      #
      # @return [Symbol, NilClass] key of the cell, returns nil if the cell doesn't belong to a table
      def key
        table.header[index].to_sym if table
      end

      def inspect
        txt = "<Workbook::Cell @value=#{value}"
        txt += " @format=#{format}" if format?
        txt += ">"
        txt
      end

      # convert value to string, and in case of a Date or Time value, apply formatting
      # @return [String]
      def to_s
        if (self.is_a? Date or self.is_a? Time) and format[:number_format]
          value.strftime(format[:number_format])
        elsif (self.class == Workbook::Cell)
          value.to_s
        else
          super
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
end
