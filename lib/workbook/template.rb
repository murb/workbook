# frozen_string_literal: true
# frozen_string_literal: true

require "workbook/modules/raw_objects_storage"

module Workbook
  # Workbook::Template is a container for different Workbook::Format's and the storage of raw template data that isn't really supported by Workbook, but should survive a typical read/write cyclus.
  class Template
    include Workbook::Modules::RawObjectsStorage

    # Initialize Workbook::Template
    def initialize
      @formats = {}
      @has_header = true
    end

    # Whether the template has a predefined header (headers are used )
    def has_header?
      @has_header
    end

    # Add a Workbook::Format to the template
    # @param [Workbook::Format] format (of a cell) to add to the template
    def add_format format
      if format.is_a? Workbook::Format
        if format.name
          @formats[format.name] = format
        else
          @formats[@formats.keys.count] = format
        end
      else
        raise ArgumentError, "format should be a Workboot::Format"
      end
    end

    # Return the list of associated formats
    # @return [Hash] A keyed-hash of named formats
    attr_reader :formats

    # Create or find a format by name
    # @return [Workbook::Format] The new or found format
    # @param [String] name of the format (e.g. whatever you want, in diff names such as 'destroyed', 'updated' and 'created' are being used)
    # @param [Symbol] variant can also be a strftime formatting string (e.g. "%Y-%m-%d")
    def create_or_find_format_by name, variant = :default
      fs = @formats[name]
      fs = @formats[name] = {} if fs.nil?
      f = fs[variant]
      if f.nil?
        @formats[name][variant] = if (variant != :default) && fs[:default]
          fs[:default].clone
        else
          Workbook::Format.new
        end
      end
      @formats[name][variant]
    end

    def set_default_formats!
      header_fmt = create_or_find_format_by :header
      header_fmt[:font_weight] = "bold"
    end
  end
end
