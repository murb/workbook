module Workbook
  class Table < Array
    attr_accessor :sheet
    
    def initialize row_cel_values=[], sheet=nil
      #@rows = []
      row_cel_values = [] if row_cel_values == nil
      row_cel_values.each do |r|
        if r.is_a? Workbook::Row
          r.table = self
        else
          r = Workbook::Row.new(r,self)
        end
      end
      self.sheet = sheet
    end
    
    def header
      first
    end
    
    # factory pattern...?
    def new_row cel_values=[]
      r = Workbook::Row.new(cel_values,self)
      return r
    end
    
  end
end
