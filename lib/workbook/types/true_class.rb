# frozen_string_literal: true
require 'workbook/cell'

module Workbook
  module Types
    class TrueClass < TrueClass
      include Workbook::Cell
    end
  end
end