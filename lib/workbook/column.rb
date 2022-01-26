# frozen_string_literal: true
# frozen_string_literal: true

module Workbook
  # Column helps us to store general properties of a column, and lets us easily perform operations on values within a column
  class Column
    attr_accessor :limit, :width # character limit

    def initialize(table = nil, options = {})
      self.table = table
      options.each { |k, v| public_send("#{k}=", v) }
    end

    # Returns column type, either :primary_key, :string, :text, :integer, :float, :decimal, :datetime, :date, :binary, :boolean
    def column_type
      return @column_type if defined?(@column_type)
      ind = index
      table[1..500].each do |row|
        if row[ind]&.cell_type
          cel_column_type = row[ind].cell_type
          if !defined?(@column_type) || @column_type.nil?
            @column_type = cel_column_type
          elsif (cel_column_type == @column_type) || (cel_column_type == :nil)
          else
            @column_type = :string
            break
          end
        end
      end
      @column_type
    end

    # Returns index of the column within the table's columns-set
    # @return [Integer, NilClass]
    def index
      table.columns.index self
    end

    # Set the table this column belongs to
    # @param [Workbook::Table] table this column belongs to
    def table= table
      raise(ArgumentError, "value should be nil or Workbook::Table") unless [NilClass, Workbook::Table].include? table.class
      @table = table
    end

    # @return [Workbook::Table]
    attr_reader :table

    def column_type= column_type
      if [:primary_key, :string, :text, :integer, :float, :decimal, :datetime, :date, :binary, :boolean].include? column_type
        @column_type = column_type
      else
        raise ArgumentError, "value should be a symbol indicating a primitive type, e.g. a string, or an integer (valid values are: :primary_key, :string, :text, :integer, :float, :decimal, :datetime, :date, :binary, :boolean)"

      end
    end

    def head_value
      table.header[index].value
    rescue
      "!noheader!"
    end

    def inspect
      "<Workbook::Column index=#{index}, header=#{head_value}>"
    end

    # default cell
    attr_reader :default

    def default= value
      @default = value if value.instance_of?(Cell)
      @default = Cell.new(value)
    end

    class << self
      # Helps to convert from e.g. "AA" to 26
      # @param [String] string that typically identifies a column
      # @return [Integer]
      def alpha_index_to_number_index string
        sum = 0
        string.upcase.chars.each_with_index do |char, char_index|
          sum = sum * 26 + char.unpack1("U") - 64
        end
        sum - 1
      end
    end
  end
end
