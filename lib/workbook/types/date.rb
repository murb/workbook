# frozen_string_literal: true

module Workbook
  module Types
    class Date < Date
      include Workbook::Modules::Cell

      def initialize(*args)
        super
      end

      def value
        self
      end

      def value= a
        puts "#value= is deprecated"
      end
    end
  end
end
