# frozen_string_literal: true

require "workbook/cell"

module Workbook
  module Types
    class Time < Time
      include Workbook::Cell
    end
  end
end
