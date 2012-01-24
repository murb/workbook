module Workbook
  class Sheet < Array
    def initialize table=Workbook::Table.new
      push table
    end
        
    def table
      first
    end
  end
end