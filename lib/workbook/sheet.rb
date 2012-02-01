module Workbook
  class Sheet < Array
    attr_accessor :book
    
    def initialize table=Workbook::Table.new([], self), book=nil
      if table.is_a? Workbook::Table
        push table
      else
        push Workbook::Table.new(table, self)
      end
      self.book = book
    end
        
    def table
      first
    end
  end
end