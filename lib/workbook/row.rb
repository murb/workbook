# frozen_string_literal: true
# frozen_string_literal: true

module Workbook
  class Row < Array
    include Workbook::Modules::Cache

    alias compare_without_header <=>
    attr_accessor :placeholder # The placeholder attribute is used in compares (corresponds to newly created or removed lines (depending which side you're on)
    attr_accessor :format

    # Initialize a new row
    #
    # @param [Workbook::Row, Array<Workbook::Cell>, Array] cells list of cells to initialize the row with, default is empty
    # @param [Workbook::Table] table a row normally belongs to a table, reference it here
    # @param [Hash] options  Supported options: parse_cells_on_batch_creation (parse cell values during row-initalization, default: false), cell_parse_options (default {}, see Workbook::Modules::TypeParser)
    def initialize cells = [], table = nil, options = {}
      options = options ? {parse_cells_on_batch_creation: false, cell_parse_options: {}, clone_cells: false}.merge(options) : {}
      add_cells(cells, options)
      self.table = table
    end

    # An internal function used in diffs
    #
    # @return [Boolean] returns true when this row is not an actual row, but a placeholder row to 'compare' against
    def placeholder?
      placeholder ? true : false
    end

    # Returns the table this row belongs to
    #
    # @return [Workbook::Table] the table this row belongs to
    def table
      @table if defined?(@table)
    end

    # Set reference to the table this row belongs to without adding the row to the table
    #
    # @param [Workbook::Table] t the table this row belongs to
    def set_table(t)
      @table = t
    end

    # Set reference to the table this row belongs to and add the row to this table
    #
    # @param [Workbook::Table] t the table this row belongs to
    def table= t
      raise ArgumentError, "table should be a Workbook::Table (you passed a #{t.class})" unless t.is_a?(Workbook::Table) || t.nil?
      if t
        @table = t
        table << self
      end
    end

    # Add cell
    # @param [Workbook::Cell, Numeric,String,Time,Date,TrueClass,FalseClass,NilClass] cell or value to add
    def push(cell)
      cell = Workbook::Cell.new(cell, {row: self}) unless cell.class == Workbook::Cell
      super(cell)
    end

    # Add cell
    # @param [Workbook::Cell, Numeric,String,Time,Date,TrueClass,FalseClass,NilClass] cell or value to add
    def <<(cell)
      cell = Workbook::Cell.new(cell, {row: self}) unless cell.class == Workbook::Cell
      super(cell)
    end

    # plus
    # @param [Workbook::Row, Array] row to add
    # @return [Workbook::Row] a new row, not linked to the table
    def +(other)
      rv = super(other)
      rv = Workbook::Row.new(rv) unless rv.class == Workbook::Row
      rv
    end

    # concat
    # @param [Workbook::Row, Array] row to add
    # @return [self] self
    def concat(row)
      row = Workbook::Row.new(row) unless row.class == Workbook::Row
      super(row)
    end

    # Overrides normal Array's []-function with support for symbols that identify a column based on the header-values and / or
    #
    # @example Lookup using fixnum or header value encoded as symbol
    #   row[1] #=> <Cell value="a">
    #   row["A"] #=> <Cell value="a">
    #   row[:a] #=> <Cell value="a">
    #
    # @param [Integer, Symbol, String] index_or_hash that identifies the column (strings are converted to symbols)
    # @return [Workbook::Cell, nil]
    def [](index_or_hash)
      if index_or_hash.is_a? Symbol
        rv = nil
        begin
          rv = to_hash[index_or_hash]
        rescue NoMethodError
        end
        rv
      elsif index_or_hash.is_a?(String) && index_or_hash.match(/^[A-Z]*$/)
        # it looks like a column indicator
        to_a[Workbook::Column.alpha_index_to_number_index(index_or_hash)]
      elsif index_or_hash.is_a? String
        symbolized = Workbook::Cell.new(index_or_hash, {row: self}).to_sym
        self[symbolized]
      elsif index_or_hash
        to_a[index_or_hash]
      end
    end

    # Overrides normal Array's []=-function with support for symbols that identify a column based on the header-values
    #
    # @example Lookup using fixnum or header value encoded as symbol (strings are converted to symbols)
    #   row[1] #=> <Cell value="a">
    #   row[:a] #=> <Cell value="a">
    #
    # @param [Integer, Symbol, String] index_or_hash that identifies the column
    # @param [String, Integer, NilClass, Date, DateTime, Time, Float] value
    # @return [Workbook::Cell, nil]
    def []= index_or_hash, value
      index = index_or_hash
      if index_or_hash.is_a? Symbol
        index = table_header_keys.index(index_or_hash)
      elsif index_or_hash.is_a?(String) && index_or_hash.match(/^[A-Z]*$/)
        # it looks like a column indicator
        index = Workbook::Column.alpha_index_to_number_index(index_or_hash)
      elsif index_or_hash.is_a? String
        symbolized = Workbook::Cell.new(index_or_hash, {row: self}).to_sym
        index = table_header_keys.index(symbolized)
      end

      value_celled = Workbook::Cell.new
      if value.is_a? Workbook::Cell
        value_celled = value
      else
        current_cell = self[index]
        if current_cell.is_a? Workbook::Cell
          value_celled = current_cell
        end
        value_celled.value = value
      end
      value_celled.row = self
      super(index, value_celled)
    end

    # Returns an array of cells allows you to find cells by a given color, normally a string containing a hex
    #
    # @param [String] color  a CSS-style hex-string
    # @param [Hash] options Option :hash_keys (default true) returns row as an array of symbols
    # @return [Array<Symbol>, Workbook::Row<Workbook::Cell>]
    def find_cells_by_background_color color = :any, options = {}
      options = {hash_keys: true}.merge(options)
      cells = collect { |c| c if c.format.has_background_color?(color) }.compact
      r = Row.new cells
      options[:hash_keys] ? r.to_symbols : r
    end

    # Returns true when the row belongs to a table and it is the header row (typically the first row)
    #
    # @return [Boolean]
    def header?
      !table.nil? && (object_id == table_header.object_id)
    end

    # Is this the first row in the table
    #
    # @return [Boolean, NilClass] returns nil if it doesn't belong to a table, false when it isn't the first row of a table and true when it is.
    def first?
      !table.nil? && (object_id == table.first.object_id)
    end

    # Returns true when all the cells in the row have values whose to_s value equals an empty string
    #
    # @return [Boolean]
    def no_values?
      all? { |c| c.value.to_s == "" }
    end

    # Converts a row to an array of symbol representations of the row content, see also: Workbook::Cell#to_sym
    # @return [Array<Symbol>] returns row as an array of symbols
    def to_symbols
      fetch_cache(:to_symbols) {
        collect { |c| c.to_sym }
      }
    end

    # Converts the row to an array of Workbook::Cell's
    # @return [Array<Workbook::Cell>] returns row as an array of symbols
    def to_a
      collect { |c| c }
    end

    def table_header
      table.header
    end

    def table_header_keys
      table_header.to_symbols
    end

    # Returns a hash representation of this row
    #
    # @return [Hash]
    def to_hash
      keys = table_header_keys
      values = self
      hash = {}
      keys.each_with_index { |k, i| hash[k] = values[i] }
      hash
    end

    # Quick assessor to the book's template, if it exists
    #
    # @return [Workbook::Template]
    def template
      table&.template
    end

    # Returns a hash representation of this row
    #
    # it differs from #to_hash as it doesn't contain the Workbook's Workbook::Cell-objects,
    # but the actual values contained in these cells
    #
    # @return [Hash]

    def to_hash_with_values
      @hash_with_values = {}
      table_header_keys.each_with_index { |k, i| @hash_with_values[k] = self[i]&.value }
      @hash_with_values
    end

    # Compares one row wiht another
    #
    # @param [Workbook::Row] other row to compare against
    # @return [Workbook::Row] a row with the diff result.
    def <=> other
      a = header? ? 0 : 1
      b = other.header? ? 0 : 1
      return (a <=> b) if (a == 0) || (b == 0)
      compare_without_header other
    end

    # The first cell of the row is considered to be the key
    #
    # @return [Workbook::Cell] the key cell
    def key
      first
    end

    # Compact detaches the row from the table
    def compact
      clone.collect { |c| c unless c.nil? }.compact
    end

    # clone the row with together with the cells
    #
    # @return [Workbook::Row] a cloned copy of self with cells
    def clone
      Workbook::Row.new(self, nil, {clone_cells: true})
    end

    # remove all the trailing nil-cells (returning a trimmed clone)
    #
    # @param [Integer] desired_length of the new row
    # @return [Workbook::Row] a trimmed clone of the array
    def trim(desired_length = nil)
      clone.trim!(desired_length)
    end

    # remove all the trailing nil-cells (returning a trimmed self)
    #
    # @param [Integer] desired_length of the new row
    # @return [Workbook::Row] self
    def trim!(desired_length = nil)
      self_count = count - 1
      count.times do |index|
        index = self_count - index
        if desired_length && (index < desired_length)
          break
        elsif desired_length && (index >= desired_length)
          delete_at(index)
        elsif self[index].nil?
          delete_at(index)
        else
          break
        end
      end
      (desired_length - count).times { |a| self << Workbook::Cell.new(nil) } if desired_length && ((desired_length - count) > 0)
      self
    end

    private

    # Only to be used on initialization
    # @param [Workbook::Row, Array<Workbook::Cell>, Array] cells list of cells to initialize the row with, default is empty
    # @param [Hash] options  Supported options: parse_cells_on_batch_creation (parse cell values during row-initalization, default: false), cell_parse_options (default {}, see Workbook::Modules::TypeParser)
    def add_cells cells = [], options = {}
      return nil if cells.nil?
      cells.each do |c|
        push(prepare_cell(c, options))
      end
    end

    def prepare_cell cell, options
      if cell.is_a? Workbook::Cell
        cell = cell.clone if options[:clone_cells]
      else
        cell = Workbook::Cell.new(cell, {row: self})
        cell.parse!(options[:cell_parse_options]) if options[:parse_cells_on_batch_creation]
      end

      cell
    end
  end
end
