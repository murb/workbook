# frozen_string_literal: true
# frozen_string_literal: true

module Workbook
  # A Sheet is a container of tables
  class Sheet
    include Enumerable
    # Initialize a new sheet
    #
    # @param [Workbook::Table, Array<Array>] table   The first table of this sheet
    # @param [Workbook::Book] book                   The book this sheet belongs to
    # @param [Hash] options                          are forwarded to Workbook::Table.new
    # @return [Workbook::Sheet] (self)
    def initialize table = Workbook::Table.new([], self), book = nil, options = {}
      @tables = []

      if table.is_a?(Workbook::Table)
        @tables.push table
      else
        @tables.push(Workbook::Table.new(table, self, options))
      end
      self.book = book
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
      @tables.first
    end

    # Returns the name of this sheet
    #
    # @return [String] the name, defaulting to sheet#index when none is set
    def name
      @name ||= "Sheet #{book.index(self) + 1}"
    end

    # Set the name of this sheet
    #
    # @param [String] the new name of the sheet
    # @return [String] the given name
    attr_writer :name

    # Set the first table of this sheet with a table or array of cells/values
    # @param [Workbook::Table, Array<Array>] table   The new first table of this sheet
    # @param [Hash] options                          are forwarded to Workbook::Table.new
    # @return [Workbook::Table] the first table of this sheet
    def table= table, options = {}
      @tables[0] = if table.is_a? Workbook::Table
        table
      else
        Workbook::Table.new(table, self, options)
      end
    end

    # Returns the Book this Sheet belongs to
    #
    # @return [Workbook::Book] the book this sheet belongs to
    def book
      @book || (self.book = Workbook::Book.new(self))
    end

    # Returns a existing or new Table at the specified index
    #
    # @return [Workbook::Table]
    def [] index
      @tables[index] ||= Workbook::Table.new([], self)
    end

    # Inserts a new Table at the specified index
    #
    # @param [Integer] index of the table
    # @param [Workbook::Table, Array<Array>] table   The new first table of this sheet
    #
    # @return [Workbook::Table]
    def []= index, value
      table_to_insert = value.is_a?(Workbook::Table) ? value : Workbook::Table.new
      table_to_insert.sheet = self
      @tables[index] = table_to_insert
    end

    # Returns an enumerator
    def each(...)
      @tables.each(...)
    end

    attr_writer :book

    # Removes all lines from this table
    #
    # @return [Workbook::Table] (self)
    def delete_all
      @tables.delete_if { |b| true }
    end

    # clones itself *and* the tables it contains
    #
    # @return [Workbook::Sheet] The cloned sheet
    def clone
      clone = super
      @tables = @tables.map(&:clone)
      clone
    end

    # Create or open the existing table at an index value
    #
    # @param [Integer] index    the index of the table
    def create_or_open_table_at index
      self[index]
    end
  end
end
