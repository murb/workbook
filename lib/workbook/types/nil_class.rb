require 'workbook/cell'

module Workbook
  module Types
    class NilClass < NilClass
      include Workbook::Cell
    end
  end
end