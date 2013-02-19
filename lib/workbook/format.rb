# -*- encoding : utf-8 -*-
require 'workbook/modules/raw_objects_storage'

module Workbook
  # Format is an object used for maintinaing a cell's formatting. It can belong to many cells. It maintains a relation to the raw template's equivalent, to preserve attributes Workbook cannot modify/access.
  class Format < Hash
    include Workbook::Modules::RawObjectsStorage
    alias_method :merge_hash, :merge
    
    attr_accessor :name
    
    # Initialize
    # @param [Workbook::Format, Hash] options (e.g. :background, :color
    def initialize options={}
      options.each {|k,v| self[k]=v}
    end
    
    # Does the current format feature a background *color*? (not black or white or transparant).
    def has_background_color? color=:any
      if self[:background_color]
        return (self[:background_color].downcase==color.to_s.downcase or (!(self[:background_color]==nil or (self[:background_color].is_a? String and (self[:background_color].downcase=='#ffffff' or self[:background_color]=='#000000'))) and color==:any))
      else
        return false
      end
    end
    
    # Returns a string that can be used as inline cell styling (e.g. `<td style="<%=cell.format.to_css%>"><%=cell%></td>`)
    # @return [String] very basic CSS styling string
    def to_css
      css_parts = []
      css_parts.push("background: #{self[:background_color].to_s} #{self[:background].to_s}".strip) if self[:background] or self[:background_color]
      css_parts.push("color: #{self[:color].to_s}") if self[:color]
      css_parts.join("; ")
    end
    
    # Combines the formatting options of one with another, removes as a consequence the reference to the raw object's equivalent.
    # @param [Workbook::Format] other_format
    def merge(other_format)
      self.remove_all_raws!
      self.merge_hash(other_format)
    end
  end
end
