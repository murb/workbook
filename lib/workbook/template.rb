require 'lib/workbook/modules/raw_objects_storage'

module Workbook
	class Template
    include Workbook::Modules::RawObjectsStorage
    
    def initialize 
      @formats = {}
      @has_header = true
    end
    
    def has_header?
      @has_header
    end
    
    def has_header= boolean
      if format.is_a? TrueClass or format.is_a? FalseClass
        @has_header = boolean
      else
        raise ArgumentError, "format should be a boolean, true of false"
      end  
    end
      
    def add_format format
      if format.is_a? Workbook::Format
        @formats[format.name]=format
      else
        raise ArgumentError, "format should be a Workboot::Format"
      end      
    end
    
    def formats
      @formats
    end
    
    def create_or_find_format_by name
      f = @formats[name]
      f = @formats[name]=Workbook::Format.new if f.nil?
      f.name = name
      f
    end
	end
end