# frozen_string_literal: true
# frozen_string_literal: true

require "workbook/modules/diff_sort"
require "workbook/writers/csv_table_writer"
require "workbook/writers/json_table_writer"
require "workbook/writers/html_writer"

module Workbook
  # A table is a container of rows and keeps track of the sheet it belongs to and which row is its header. Additionally suport for CSV writing and diffing with another table is included.
  class Table
    include Enumerable
    extend Forwardable

    include Workbook::Modules::TableDiffSort
    include Workbook::Writers::CsvTableWriter
    include Workbook::Writers::JsonTableWriter
    include Workbook::Writers::HtmlTableWriter

    delegate [:first, :last, :pop, :delete_at, :each, :collect, :each_with_index, :count, :index, :delete_if, :include?] => :@rows

    attr_accessor :name
    attr_reader :rows

    def initialize row_cel_values = [], sheet = nil, options = {}
      @rows = []
      row_cel_values = [] if row_cel_values.nil?
      row_cel_values.each_with_index do |r, ri|
        if r.is_a? Workbook::Row
          r.table = self
        else
          r = Workbook::Row.new(r, self, options)
        end
        define_columns_with_row(r) if ri == 0
      end
      self.sheet = sheet
      # Column data is considered as a 'row' with 'cells' that contain 'formatting'
    end

    # Quick assessor to the book's template, if it exists
    #
    # @return [Workbook::Template]
    def template
      sheet.book.template
    end

    # Returns the header of this table (typically the first row, but can be a different row).
    # The header row is also used for finding values in a aribrary row.
    #
    # @return [Workbook::Row] The header
    def header
      if defined?(@header) && (@header == false)
        false
      elsif defined?(@header) && @header
        @header
      else
        first
      end
    end

    # Set the header of this table (typically the first row, but can be a different row).
    # The header row is also used for finding values in a aribrary row.
    #
    # @param [Workbook::Row, Integer] h should be the row or the index of this table's row
    # @return [Workbook::Row] The header
    def header= h
      @header = if h.is_a? Numeric
        self[h]
      else
        h
      end
    end

    # Returns the index of the header row
    #
    # @return [Integer] The index of the header row (typically 0)
    def header_row_index
      @rows.map(&:object_id).index(header.object_id)
    end

    def define_columns_with_row(r)
      self.columns = r.collect do |column|
        Workbook::Column.new self, {}
      end
    end

    # Generates a new row, with optionally predefined cell-values, that is already connected to this table.
    #
    # @param [Array, Workbook::Row] cell_values is an array or row of cell values
    # @return [Workbook::Row] the newly created row
    def new_row cell_values = []
      Workbook::Row.new(cell_values, self)
    end

    def create_or_open_row_at index
      r = self[index]
      if r.nil?
        r = Workbook::Row.new
        r.table = self
      end
      r
    end

    # Removes all empty lines. This function is particularly useful if you typically add lines to the end of a template-table, which sometimes has unremovable empty lines.
    #
    # @return [Workbook::Table] self
    def remove_empty_lines!
      delete_if { |r| r.nil? || r.compact.empty? }
      self
    end

    # Add row
    # @param [Workbook::Table, Array] row(s) to add
    def push(*args)
      args.each do |arg|
        self.<<(arg)
      end
    end
    alias_method :append, :push

    # Add row
    # @param [Workbook::Table, Array] row to add
    def <<(row)
      row = row.is_a?(Workbook::Row) ? row : Workbook::Row.new(row)
      row.set_table(self)
      @rows << row
    end

    # Insert row
    # @param [Integer] index where to add
    # @param [Array, Workbook::Row] row(s) to add
    def insert index, *rows
      rows.each_with_index do |row, i|
        row = row.is_a?(Workbook::Row) ? row : Workbook::Row.new(row)
        row.set_table(self)
        row.insert(index + i, row)
      end
    end

    def has_contents?
      clone.remove_empty_lines!.count != 0
    end

    # Returns true if the row exists in this table
    #
    # @param [Workbook::Row] row to test for
    # @return [Boolean] whether the row exist in this table
    def contains_row? row
      raise ArgumentError, "table should be a Workbook::Row (you passed a #{t.class})" unless row.is_a?(Workbook::Row)
      collect { |r| r.object_id }.include? row.object_id
    end

    # Returns the sheet this table belongs to, creates a new sheet if none exists
    #
    # @return [Workbook::Sheet] The sheet this table belongs to
    def sheet
      return @sheet if defined?(@sheet) && !@sheet.nil?
      self.sheet = Workbook::Sheet.new(self)
    end

    # Returns the sheet this table belongs to, creates a new sheet if none exists
    #
    # @param [Workbook::Sheet] sheet this table belongs to
    # @return [Workbook::Sheet] The sheet this table belongs to
    attr_writer :sheet

    # Removes all lines from this table
    #
    # @return [Workbook::Table] (self)
    def delete_all
      delete_if { |b| true }
    end

    # clones itself *and* the rows it contains
    #
    # @return [Workbook::Table] The cloned table
    def clone
      table_clone = super
      @rows = @rows.map { |row|
        cloned_row = row.clone
        cloned_row.set_table(self)
        cloned_row
      }
      self.header = @rows[header_row_index] if header_row_index
      table_clone.header = table_clone.rows[header_row_index] if header_row_index
      table_clone
    end

    # Overrides normal Array's []-function with support for symbols that identify a column based on the header-values
    #
    # @example Lookup using fixnum or header value encoded as symbol
    #   table[0] #=> <Row [a,2,3,4]> (first row)
    #   table["A1"] #=> <Cell value="a"> (first cell of first row)
    #
    # @param [Integer, String] index_or_string to reference to either the row, or the cell
    # @return [Workbook::Row, Workbook::Cell, nil]
    def [] index_or_string
      if index_or_string.is_a? String
        match = index_or_string.match(/([A-Z]+)([0-9]*)/i)
        col_index = Workbook::Column.alpha_index_to_number_index(match[1])
        if match[2] == ""
          columns[col_index]
        else
          row_index = match[2].to_i - 1
          @rows[row_index][col_index]
        end
      elsif index_or_string.is_a? Symbol
        header[index_or_string].column
      elsif index_or_string.is_a? Range
        collection = @rows[index_or_string].collect { |a| a.clone }
        Workbook::Table.new collection
      elsif index_or_string.is_a? Integer
        @rows[index_or_string]
      end
    end

    # Overrides normal Row's []=-function; automatically converting to row and setting
    # with the label correctly
    #
    # @example Lookup using fixnum or header value encoded as symbol
    #   `table[0] = <Row [a,2,3,4]>` (set first row)
    #   `table["A1"] = 2` (set first cell of first row to 2)
    #
    # @param [Integer, String] index_or_string to reference to either the row, or the cell
    # @param [Workbook::Table, Array] new_value to set
    # @return [Workbook::Cell, nil]
    def []= index_or_string, new_value
      if index_or_string.is_a? String
        match = index_or_string.upcase.match(/([A-Z]*)([0-9]*)/)
        cell_index = Workbook::Column.alpha_index_to_number_index(match[1])
        row_index = match[2].to_i - 1
        @rows[row_index][cell_index].value = new_value
      else
        row = new_value.is_a?(Workbook::Row) ? new_value : Workbook::Row.new(new_value)
        row.table = self
      end
    end

    def == other
      to_csv == other.to_csv
    end

    # remove all the trailing empty-rows (returning a trimmed clone)
    #
    # @param [Integer] desired_row_length of the rows
    # @return [Workbook::Row] a trimmed clone of the array
    def trim desired_row_length = nil
      clone.trim!(desired_row_length)
    end

    # remove all the trailing empty-rows (returning a trimmed self)
    #
    # @param [Integer] desired_row_length of the new row
    # @return [Workbook::Row] self
    def trim! desired_row_length = nil
      max_length = collect { |a| a.trim.length }.max
      self_count = count - 1
      count.times do |index|
        index = self_count - index
        if @rows[index].trim.empty?
          delete_at(index)
        else
          break
        end
      end
      each { |a| a.trim!(max_length) }
      self
    end

    # Returns The dimensions of this sheet based on longest row
    # @return [Array] x-width, y-height
    def dimensions
      height = count
      width = collect { |a| a.length }.max
      [width, height]
    end

    # Returns an array of Column-classes describing the columns of this table
    # @return [Array<Column>] columns
    def columns
      @columns ||= header.collect { |header_cell|
        Column.new(self)
      }
    end

    # Returns an array of Column-classes describing the columns of this table
    # @param [Array<Column>] columns
    # @return [Array<Column>] columns
    def columns= columns
      columns.each { |c| c.table = self }
      @columns = columns
    end
  end
end
