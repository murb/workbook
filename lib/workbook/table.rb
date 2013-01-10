require 'workbook/modules/table_diff_sort'
require 'workbook/writers/csv_table_writer'


module Workbook  
  class Table < Array
    include Workbook::Modules::TableDiffSort
    include Workbook::Writers::CsvTableWriter
    attr_accessor :sheet
    attr_accessor :header
    
    def initialize row_cel_values=[], sheet=nil, options={}
      #@rows = []
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
    
    def contains_row? row
      raise ArgumentError, "table should be a Workbook::Row (you passed a #{t.class})" unless row.is_a?(Workbook::Row)
      self.collect{|r| r.object_id}.include? row.object_id
    end
    
  end
end
