require 'workbook/cell'

module Workbook
  module Types
    class Date < Date
      include Workbook::Cell
    end
  end
end