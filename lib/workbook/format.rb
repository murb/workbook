module Workbook
  class Format < Hash
    attr_accessor :raw
    
    def initialize options={}
      options.each {|k,v| self[k]=v}
      @raw = options.raw if options.methods.include? "raw"
    end
    
  end
end
