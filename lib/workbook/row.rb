module Workbook
  class Row < Array
    alias_method :compare_without_header, :<=>
    
    # The placeholder attribute is used in compares (corresponds to newly created or removed lines (depending which side you're on)
    attr_accessor :placeholder
    attr_accessor :format
    
    
    # @param [Workbook::Row, Array<Workbook::Cell>, Array] list of cells to initialize the row with, default is empty
    # @param [Workbook::Table] a row normally belongs to a table, reference it here
    # @param [Hash], option hash. Supprted options: parse_cells_on_batch_creation (parse cell values during row-initalization, default: false), cell_parse_options (default {}, see Workbook::Modules::TypeParser)
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
    
    def placeholder?
      placeholder ? true : false
    end
    
    def table
      @table
    end
    
    # Set reference to the table this row belongs to (and adds the row to this table)
    #
    # @param [Workbook::Table] 
    def table= t
      raise ArgumentError, "table should be a Workbook::Table (you passed a #{t.class})" unless t.is_a?(Workbook::Table) or t == nil
      if t
        @table = t
        table << self
      end
    end
    
    # Overrides normal Array's []-function with support for symbols that identify a column based on the header-values
    #
    # @example Lookup using fixnum or header value encoded as symbol
    #   row[1] #=> <Cell value="a">
    #   row[:a] #=> <Cell value="a">
    #
    # @param [Fixnum, Symbol]
    # @return [Workbook::Cell, nil]
    def [](index_or_hash)
      if index_or_hash.is_a? Fixnum
        return to_a[index_or_hash]
      elsif index_or_hash.is_a? Symbol
        rv = nil
        begin
          rv = to_hash[index_or_hash]
        rescue NoMethodError
        end
        return rv
      end
    end
    
    # Returns an array of cells allows you to find cells by a given color, normally a string containing a hex
    #
    # @param [String] default :any colour, can be a CSS-style hex-string
    # @param [Hash] options. Option :hash_keys (default true) returns row as an array of symbols
    # @returns [Array<Symbol>, Workbook::Row<Workbook::Cell>]
    def find_cells_by_background_color color=:any, options={}
      options = {:hash_keys=>true}.merge(options)
      cells = self.collect {|c| c if c.format.has_background_color?(color) }.compact
      r = Row.new cells
      options[:hash_keys] ? r.to_symbols : r
    end
    
    # Returns true when the row belongs to a table and it is the header row (typically the first row) 
    #
    # @returns [@Boolean]
    def header?
      table != nil and self.object_id == table.header.object_id
    end
    
    # @returns [Array<Symbol>] returns row as an array of symbols
    def to_symbols
      collect{|c| c.to_sym}
    end
    
    # @returns [Array<Workbook::Cell>] returns row as an array of symbols
    def to_a
      self.collect{|c| c}
    end
    
    def to_hash
      return @hash if @hash
      keys = table.header.to_symbols
      values = self
      @hash = {}
      keys.each_with_index {|k,i| @hash[k]=values[i]}
      return @hash
    end
    
    def <=> other
      a = self.header? ? 0 : 1
      b = other.header? ? 0 : 1
      return (a <=> b) if (a==0 or b==0)
      compare_without_header other
    end
    
    def key
      first
    end
    
    # Compact detaches the row from the table
    def compact
      r = self.clone
      r = r.collect{|c| c unless c.nil?}.compact
    end
  end
end
