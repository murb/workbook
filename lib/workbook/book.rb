module Workbook
  class Book < Array
    def initialize sheet=Workbook::Sheet.new
      push sheet if sheet
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
