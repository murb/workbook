# frozen_string_literal: true

module Workbook
  module Types
    class FalseClass < FalseClass
      include Workbook::Modules::Cell
    end
  end
end
