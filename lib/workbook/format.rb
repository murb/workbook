module Workbook
  class Format < Hash
    attr_accessor :raw
    attr_accessor :name
    
    def initialize options={}
      options.each {|k,v| self[k]=v}
      @raw = options.raw if options.methods.include? "raw"
    end
    
  end
end
