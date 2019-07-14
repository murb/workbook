# frozen_string_literal: true
# frozen_string_literal: true

require "workbook/modules/cell"

module Workbook
  class Cell
    include Workbook::Modules::Cell

    # @param [Numeric,String,Time,Date,TrueClass,FalseClass,NilClass] value a valid value
    # @param [Hash] options a reference to :format (Workbook::Format) can be specified
    def initialize value = nil, options = {}
      self.format = options[:format] if options[:format]
      self.row = options[:row]
      self.value = value
      @to_sym = nil
    end
  end
end
