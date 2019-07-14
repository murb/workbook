# frozen_string_literal: true

require "workbook/cell"

module Workbook
  module Types
    class Numeric < Numeric
      include Workbook::Cell
    end
  end
end
