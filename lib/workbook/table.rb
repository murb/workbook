require 'workbook/modules/table_diff_sort'
require 'workbook/writers/csv_table_writer'


module Workbook  
  # A table is a container of rows and keeps track of the sheet it belongs to and which row is its header. Additionally suport for CSV writing and diffing with another table is included.
  class Table < Array
    include Workbook::Modules::TableDiffSort
    include Workbook::Writers::CsvTableWriter
    attr_accessor :sheet
    attr_accessor :header
    
    def initialize row_cel_values=[], sheet=nil, options={}
      row_cel_values = [] if row_cel_values == nil
      row_cel_values.each do |r|
        if r.is_a? Workbook::Row
          r.table = self
        else
          r = Workbook::Row.new(r,self, options)
        end
      end
      self.sheet = sheet
      # Column data is considered as a 'row' with 'cells' that contain 'formatting'
    end
    
    # Returns the header of this table (typically the first row, but can be a different row). 
    # The header row is also used for finding values in a aribrary row.
    #
    # @return [Workbook::Row] The header
    def header
      if @header == false
        false
      elsif @header
        @header
      else
        first
      end
    end
    
    # Generates a new row, with optionally predefined cell-values, that is already connected to this table.
    def new_row cell_values=[]
      r = Workbook::Row.new(cell_values,self)
      return r
    end
    
    def create_or_open_row_at index
      r = self[index]
      if r == nil
        r = Workbook::Row.new
        r.table=(self)
      end
      r 
    end  
    
    def remove_empty_lines!
      self.delete_if{|r| r.nil? or r.compact.empty?}
      self
    end
    
    def push(row)
      super(row)
      row.set_table(self)
    end
    
    def <<(row)
      super(row)
      row.set_table(self)
    end
    
    def has_contents?
      self.clone.remove_empty_lines!.count != 0
    end
    
    # Returns true if the row exists in this table
    #
    # @param [Workbook::Row] row to test for
    # @return [Boolean] whether the row exist in this table 
    def contains_row? row
      raise ArgumentError, "table should be a Workbook::Row (you passed a #{t.class})" unless row.is_a?(Workbook::Row)
      self.collect{|r| r.object_id}.include? row.object_id
    end
    
    # Returns the sheet this table belongs to, creates a new sheet if none exists
    #
    # @return [Workbook::Sheet] The sheet this table belongs to
    def sheet
      if @sheet
        return @sheet
      else
        self.sheet = Workbook::Sheet.new(self)
        return @sheet
      end
    end
    
    # Removes all lines from this table
    #
    # @return [Workbook::Table] (self)
    def delete_all
      self.delete_if{|b| true}
    end
    
    # clones itself *and* the rows it contains
    #
    # @return [Workbook::Table] The cloned table
    def clone
      t = self
      c = super
      header_row_index = t.index(t.header)
      c.delete_all
      t.each{|r| c << r.clone}
      c.header = c[header_row_index] if header_row_index
      return c
    end
    
  end
end
