require 'lib/workbook/modules/raw_objects_storage'

module Workbook
  class Format < Hash
    include Workbook::Modules::RawObjectsStorage
    alias_method :merge_hash, :merge
    
    attr_accessor :name
    
    def initialize options={}
      options.each {|k,v| self[k]=v}
    end
    
    def merge(a)
      self.remove_all_raws!
      self.merge_hash(a)
    end
  end
end
