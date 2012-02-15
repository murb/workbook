module Workbook
  class Row < Array
    alias_method :compare_without_header, :<=>
    
    #used in compares
    attr_accessor :placeholder
    
    
    def initialize cells=[], table=nil
      cells = [] if cells==nil
      self.table= table
      cells.each do |c| 
        if c.is_a? Workbook::Cell
          c = c
        else
          c = Workbook::Cell.new(c)
        end
        push c
      end
    end
    
    def table
      @table
    end
    
    def table= t
      
      raise ArgumentError, "table should be a Workbook::Table (you passed a #{t.class})" unless t.is_a?(Workbook::Table) or t == nil
      if t
        @table = t
        table << self unless table.collect{|r| r.object_id}.include? self.object_id
      end
    end
    
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
    
    def header?
      table != nil and self.object_id == table.header.object_id
    end
    
    def to_symbols
      collect{|c| c.to_sym}
    end
    
    def to_a
      self.collect{|c| c}
    end
    
    def to_hash
      keys = table.header.to_symbols
      values = self
      hash = {}
      keys.each_with_index {|k,i| hash[k]=values[i]}
      return hash
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
  end
end
