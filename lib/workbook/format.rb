require 'lib/workbook/modules/raw_objects_storage'

module Workbook
  class Format < Hash
    include Workbook::Modules::RawObjectsStorage
    
    attr_accessor :name
    
    def initialize options={}
      options.each {|k,v| self[k]=v}
    end
    
  end
end
