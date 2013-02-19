# -*- encoding : utf-8 -*-
module Workbook
  class Sheet < Array
    # A Sheet is a container of tables
    attr_accessor :book
    attr_accessor :name
    
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
    
    # Removes all lines from this table
    #
    # @return [Workbook::Table] (self)
    def delete_all
      self.delete_if{|b| true}
    end
    
    # clones itself *and* the tables it contains
    #
    # @return [Workbook::Sheet] The cloned sheet
    def clone
      s = self
      c = super
      c.delete_all
      s.each{|t| c << t.clone}
      return c
    end
  end
end
