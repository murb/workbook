require 'workbook/modules/table_diff_sort'
require 'workbook/writers/csv_table_writer'


module Workbook  
  class Table < Array
    include Workbook::Modules::TableDiffSort
    include Workbook::Writers::CsvTableWriter
    attr_accessor :sheet
    attr_accessor :header
    
    def initialize row_cel_values=[], sheet=nil
      #@rows = []
      row_cel_values = [] if row_cel_values == nil
      row_cel_values.each do |r|
        if r.is_a? Workbook::Row
          r.table = self
        else
          r = Workbook::Row.new(r,self)
        end
      end
      self.sheet = sheet
    end
    
    def header
      if @header == false
        false
      elsif @header
        @header
      else
        first
      end
    end
    
    # factory pattern...?
    def new_row cel_values=[]
      r = Workbook::Row.new(cel_values,self)
      return r
    end
    
    def create_or_open_row_at index
      s = self[index]
      s = self[index] = Workbook::Row.new if s == nil
      s.table = self
      s 
    end  
  end
end
