# -*- encoding : utf-8 -*-
module Workbook

  # Column helps us to store general properties of a column, and lets us easily perform operations on values within a column
  class Column
    attr_accessor :column_type # :primary_key, :string, :text, :integer, :float, :decimal, :datetime, :date, :binary, :boolean.
    attr_accessor :limit #character limit
    attr_accessor :default #default cell

    def initialize(options={})
      options.each{ |k,v| self.public_send("#{k}=",v) }
    end

    def column_type= column_type
      if [:primary_key, :string, :text, :integer, :float, :decimal, :datetime, :date, :binary, :boolean].include? column_type
        @column_type = column_type
      else
        raise ArgumentError, "value should be a symbol indicating a primitive type, e.g. a string, or an integer (valid values are: :primary_key, :string, :text, :integer, :float, :decimal, :datetime, :date, :binary, :boolean)"

      end
    end

    def default= value
      @default = value if value.class == Cell
      @default = Cell.new(value)
    end


  end
end