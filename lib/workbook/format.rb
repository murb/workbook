require 'workbook/modules/raw_objects_storage'

module Workbook
  class Format < Hash
    include Workbook::Modules::RawObjectsStorage
    alias_method :merge_hash, :merge
    
    attr_accessor :name
    
    def initialize options={}
      options.each {|k,v| self[k]=v}
    end
    
    def has_background_color? color=:any
      if self[:background_color]
        return (self[:background_color].downcase==color.to_s.downcase or (!(self[:background_color]==nil or self[:background_color]=='#ffffff' or self[:background_color]=='#000000') and color==:any))
      else
        return false
      end
    end
    
    def merge(a)
      self.remove_all_raws!
      self.merge_hash(a)
    end
  end
end
