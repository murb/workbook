module Workbook
  class Sheet < Array
    # A Sheet is a container of tables
    attr_accessor :book
    
    # Initialize a new sheet
    #
    # @param [Workbook::Table, Array<Array>] table   The first table of this sheet
    # @param [Workbook::Book] book                   The book this sheet belongs to
    # @param [Hash] options                          are forwarded to Workbook::Table.new
    # @return [Workbook::Sheet] (self)
    def initialize table=Workbook::Table.new([], self), book=nil, options={}
      if table.is_a? Workbook::Table
        push table
      else
        push Workbook::Table.new(table, self, options)
      end
      self.book = book
      return self
    end
    
    # Returns true if the first table of this sheet contains anything
    #
    # @return [Boolean]
    def has_contents?
      table.has_contents?
    end
        
    # Returns the first table of this sheet
    #
    # @return [Workbook::Table] the first table of this sheet
    def table
      first
    end
    
    # Returns the book this sheet belongs to
    #
    # @return [Workbook::Book] the book this sheet belongs to
    def book
      if @book
        return @book
      else
        self.book = Workbook::Book.new(self)
        return @book
      end
    end
  end
end