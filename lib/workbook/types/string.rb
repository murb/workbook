# frozen_string_literal: true
require 'workbook/cell'

module Workbook
  module Types
    class String < String
      include Workbook::Cell
    end
  end
end