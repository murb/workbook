module Workbook
  module Types
    class Date < Date
      include Workbook::Modules::Cell

      def initialize(*args)
        super(*args)
      end

      def value
        self
      end

      def value= a
        throw "#value= is no longer available"
      end
    end
  end
end