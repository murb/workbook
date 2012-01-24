module Workbook
  class Table < Array
    
    def initialize row_cel_values=[]
      #@rows = []
      row_cel_values.each do |r|
        if r.class == Workbook::Row
          r.table = self
        else
          r = Workbook::Row.new(r,self)
        end
      end
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
