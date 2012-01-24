module Workbook
  class Sheet < Array
    def initialize table=Workbook::Table.new
      if table.is_a? Workbook::Table
        push table
      else
        push Workbook::Table.new(table)
      end
    end
        
    def table
      first
    end
  end
end