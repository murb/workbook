# -*- encoding : utf-8 -*-
module Workbook

  # Column helps us to store general properties of a column, and lets us easily perform operations on values within a column
  class Column
    attr_accessor :limit, :table #character limit

    def initialize(table=nil, options={})
      self.table = table
      options.each{ |k,v| self.public_send("#{k}=",v) }
    end

    # Returns column type, either :primary_key, :string, :text, :integer, :float, :decimal, :datetime, :date, :binary, :boolean
    def column_type
      @column_type
    end

    def table= t
      raise(ArgumentError, "value should be nil or Workbook::Table") unless [NilClass,Workbook::Table].include? t.class
      @table = t
    end

    def column_type= column_type
      if [:primary_key, :string, :text, :integer, :float, :decimal, :datetime, :date, :binary, :boolean].include? column_type
        @column_type = column_type
      else
        raise ArgumentError, "value should be a symbol indicating a primitive type, e.g. a string, or an integer (valid values are: :primary_key, :string, :text, :integer, :float, :decimal, :datetime, :date, :binary, :boolean)"

      end
    end

    #default cell
    def default
      return @default
    end

    def default= value
      @default = value if value.class == Cell
      @default = Cell.new(value)
    end


  end
end