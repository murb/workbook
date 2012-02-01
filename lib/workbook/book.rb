require 'lib/workbook/writers/xls_writer'

module Workbook
  class Book < Array
    include Workbook::Writers::XlsWriter
    
    attr_accessor :raw
    attr_accessor :title
    
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
    
  end
end
