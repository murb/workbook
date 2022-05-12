# frozen_string_literal: true
# frozen_string_literal: true

require "workbook/modules/cell"

module Workbook
  class Cell
    include Workbook::Modules::Cell

    class << self

      # returns a symbol representation of a string.
      # @param [String] value to convert
      # @example
      #
      #     <Workbook::Cell value="yet another value">.to_sym # returns :yet_another_value
      def value_to_sym value
        v = nil
        cell_type ||= ::Workbook::Modules::Cell::CLASS_CELLTYPE_MAPPING[value.class.to_s]
        value_to_s = value.to_s.strip.downcase
        unless value.nil? || value_to_s == ""
          if cell_type == :integer
            v = "num#{value}".to_sym
          elsif cell_type == :float
            v = "num#{value}".sub(".", "_").to_sym
          else
            v = value_to_s.strip
            ends_with_exclamationmark = (v[-1] == "!")
            ends_with_questionmark = (v[-1] == "?")

            v = _replace_possibly_problematic_characters_from_string(v)

            v = v.encode(Encoding.find("ASCII"), invalid: :replace, undef: :replace, replace: "")

            v = "#{v}!" if ends_with_exclamationmark
            v = "#{v}?" if ends_with_questionmark
            v = v.downcase.to_sym
          end
        end
        v
      end

      private

      def _replace_possibly_problematic_characters_from_string(string)
        Workbook::Modules::Cell::CHARACTER_REPACEMENTS.each do |ac, rep|
          ac.each do |s|
            string = string.gsub(s, rep)
          end
        end
        string
      end

    end

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
