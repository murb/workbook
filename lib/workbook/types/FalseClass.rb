require 'workbook/cell'

module Workbook
  module Types
    class FalseClass < FalseClass
      include Workbook::Cell
    end
  end
end