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
    
    # Returns a string that can be used as inline cell styling (e.g. `<td style="<%=cell.format.to_css%>"><%=cell%></td>`)
    def to_css
      css_parts = []
      css_parts.push("background: #{self[:background_color].to_s} #{self[:background].to_s}".strip) if self[:background] or self[:background_color]
      css_parts.push("color: #{self[:color].to_s}") if self[:color]
      css_parts.join("; ")
    end
    
    def merge(a)
      self.remove_all_raws!
      self.merge_hash(a)
    end
  end
end
