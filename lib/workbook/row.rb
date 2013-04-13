# -*- encoding : utf-8 -*-
module Workbook
  class Row < Array
    alias_method :compare_without_header, :<=>
    attr_accessor :placeholder     # The placeholder attribute is used in compares (corresponds to newly created or removed lines (depending which side you're on)
    attr_accessor :format
    
    # Initialize a new row
    #
    # @param [Workbook::Row, Array<Workbook::Cell>, Array] cells list of cells to initialize the row with, default is empty
    # @param [Workbook::Table] table a row normally belongs to a table, reference it here
    # @param [Hash] options  Supprted options: parse_cells_on_batch_creation (parse cell values during row-initalization, default: false), cell_parse_options (default {}, see Workbook::Modules::TypeParser)
    def initialize cells=[], table=nil, options={}
      options=options ? {:parse_cells_on_batch_creation=>false,:cell_parse_options=>{}}.merge(options) : {} 
      cells = [] if cells==nil
      self.table= table
      cells.each do |c| 
        if c.is_a? Workbook::Cell
          c = c
        else
          c = Workbook::Cell.new(c)
          c.parse!(options[:cell_parse_options]) if options[:parse_cells_on_batch_creation]
        end
        push c
      end
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
      @table
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
      raise ArgumentError, "table should be a Workbook::Table (you passed a #{t.class})" unless t.is_a?(Workbook::Table) or t == nil
      if t
        @table = t
        table.push(self) #unless table.index(self) and self.placeholder?
      end
    end
    
    # Overrides normal Array's []-function with support for symbols that identify a column based on the header-values
    #
    # @example Lookup using fixnum or header value encoded as symbol
    #   row[1] #=> <Cell value="a">
    #   row[:a] #=> <Cell value="a">
    #
    # @param [Fixnum, Symbol] index_or_hash
    # @return [Workbook::Cell, nil]
    def [](index_or_hash)
      if index_or_hash.is_a? Symbol
        rv = nil
        begin
          rv = to_hash[index_or_hash]
        rescue NoMethodError
        end
        return rv
      else 
        if index_or_hash
          return to_a[index_or_hash]
        end
      end
    end

    # Overrides normal Array's []=-function with support for symbols that identify a column based on the header-values
    #
    # @example Lookup using fixnum or header value encoded as symbol
    #   row[1] #=> <Cell value="a">
    #   row[:a] #=> <Cell value="a">
    #
    # @param [Fixnum, Symbol] index_or_hash
    # @param [String, Fixnum, NilClass, Date, DateTime, Time, Float] value 
    # @return [Workbook::Cell, nil]
    def []= (index_or_hash, value)
      index = index_or_hash
      if index_or_hash.is_a? Symbol
        index = table_header_keys.index(index_or_hash)
      end
    
      value_celled = Workbook::Cell.new
      if value.is_a? Workbook::Cell
        value_celled = value
      else
        current_cell = self[index]
        if current_cell.is_a? Workbook::Cell
          value_celled = current_cell 
        end
        value_celled.value=(value)
      end
      super(index,value_celled)
    end
        
    # Returns an array of cells allows you to find cells by a given color, normally a string containing a hex
    #
    # @param [String] color  a CSS-style hex-string
    # @param [Hash] options Option :hash_keys (default true) returns row as an array of symbols
    # @return [Array<Symbol>, Workbook::Row<Workbook::Cell>]
    def find_cells_by_background_color color=:any, options={}
      options = {:hash_keys=>true}.merge(options)
      cells = self.collect {|c| c if c.format.has_background_color?(color) }.compact
      r = Row.new cells
      options[:hash_keys] ? r.to_symbols : r
    end
    
    # Returns true when the row belongs to a table and it is the header row (typically the first row) 
    #
    # @return [Boolean]
    def header?
      table != nil and self.object_id == table.header.object_id
    end
    
    # Is this the first row in the table
    # 
    # @return [Boolean, NilClass] returns nil if it doesn't belong to a table, false when it isn't the first row of a table and true when it is.
    def first?
      table != nil and self.object_id == table.first.object_id
    end
    
    # Returns true when all the cells in the row have values whose to_s value equals an empty string
    #
    # @return [Boolean]
    def no_values?
      all? {|c| c.value.to_s == ''}
    end
    
    # Converts a row to an array of symbol representations of the row content, see also: Workbook::Cell#to_sym
    # @return [Array<Symbol>] returns row as an array of symbols
    def to_symbols
      collect{|c| c.to_sym}
    end
    
    # Converts the row to an array of Workbook::Cell's
    # @return [Array<Workbook::Cell>] returns row as an array of symbols
    def to_a
      self.collect{|c| c}
    end
    
    def table_header_keys
      table.header.to_symbols
    end
    
    # Returns a hash representation of this row
    #
    # @return [Hash]
    def to_hash
      return @hash if @hash
      keys = table_header_keys
      values = self
      @hash = {}
      keys.each_with_index {|k,i| @hash[k]=values[i]}
      return @hash
    end
    
    # Compares one row wiht another
    #
    # @param [Workbook::Row] other row to compare against
    # @return [Workbook::Row] a row with the diff result. 
    def <=> other
      a = self.header? ? 0 : 1
      b = other.header? ? 0 : 1
      return (a <=> b) if (a==0 or b==0)
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
      r = self.clone
      r = r.collect{|c| c unless c.nil?}.compact
    end
    
    # clone the row with together with the cells
    #
    # @return [Workbook::Row] a cloned copy of self with cells
    def clone
      Workbook::Row.new(to_a.collect{|c| c.clone})
    end
  end
end
