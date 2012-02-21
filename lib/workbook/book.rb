require 'workbook/writers/xls_writer'
require 'workbook/readers/xls_reader'

module Workbook
  class Book < Array
    include Workbook::Writers::XlsWriter
    include Workbook::Readers::XlsReader
    
    attr_accessor :title
    attr_accessor :template
    
    # def initialize sheet=Workbook::Sheet.new
    #   push sheet if sheet
    # end
    
    def initialize sheet=Workbook::Sheet.new([], self)
      if sheet.is_a? Workbook::Sheet
        push sheet
      else
        push Workbook::Sheet.new(sheet, self)
      end
    end
    
    def template
      @template ||= Workbook::Template.new
    end
    
    def template= template
      raise ArgumentError, "format should be a Workboot::Format" unless template.is_a? Workbook::Template
      @template = template
    end
    
    def title
      @title ? @title : "untitled document"
    end
    
  	def push sheet=Workbook::Sheet.new
      super(sheet)
    end
    def sheet
      push Workbook::Sheet.new unless first
      first
    end
    
    # Loads an external file into an existing worbook
    def open filename, ext=nil
      f = File.open(filename,'rb')
      ext = File.extname(filename).gsub('.','') unless ext
      send("load_#{ext}".to_sym,f)
    end
    
    # Create an instance from a file, using open.
    def self.open filename, ext=nil
      wb = self.new
      wb.open filename, ext
      return wb
    end

    def create_or_open_sheet_at index
      s = self[index]
      s = self[index] = Workbook::Sheet.new if s == nil
      s.book = self
      s 
    end
    
    def sort
      raise Exception("Books can't be sorted")
    end
  end
end
