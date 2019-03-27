# frozen_string_literal: true

# -*- encoding : utf-8 -*-
# frozen_string_literal: true
require 'workbook/modules/raw_objects_storage'

module Workbook
  # Format is an object used for maintinaing a cell's formatting. It can belong to many cells. It maintains a relation to the raw template's equivalent, to preserve attributes Workbook cannot modify/access.
  # The keys in the Hash are intended to closely mimick the CSS-style options:
  #
  #     {
  #       background_color: '#ff000',
  #       color: '#ffff00',
  #       font_weight: :bold,
  #       text_decoration: :underline,
  #     }
  #
  # Note that as we speak, not all exporters support all properties properly. Consider it WIP.
  class Format < Hash
    include Workbook::Modules::RawObjectsStorage
    alias_method :merge_hash, :merge
    alias_method :merge_hash!, :merge!
    attr_accessor :name, :parent

    # Initializes Workbook::Format with a hash. The keys in the Hash are intended to closely mimick the CSS-style options (see above)
    #
    # @param [Workbook::Format, Hash] options (e.g. :background, :color, :background_color, :font_weight (integer or css-type labels)
    # @return [String] the name of the format, default: nil
    def initialize options={}, name=nil
      if options.is_a? String
        name = options
      else
        options.each {|k,v| self[k]=v}
      end
      self.name = name
    end

    # Does the current format feature a background *color*? (not black or white or transparant).
    def has_background_color? color=:any
      bg_color = flattened[:background_color] ? flattened[:background_color].to_s.downcase : nil

      if color != :any and bg_color
        return bg_color == color.to_s.downcase
      elsif bg_color
        return !(flattened[:background_color].downcase=='#ffffff' or flattened[:background_color]=='#000000')
      else
        return false
      end
    end

    # Returns a string that can be used as inline cell styling (e.g. `<td style="<%=cell.format.to_css%>"><%=cell%></td>`)
    # @return String very basic CSS styling string
    def to_css
      css_parts = []
      background = [flattened[:background_color].to_s,flattened[:background].to_s].join(" ").strip
      css_parts.push("background: #{background}") if background and background != ""
      css_parts.push("color: #{flattened[:color].to_s}") if flattened[:color]
      css_parts.join("; ")
    end

    # Combines the formatting options of one with another, removes as a consequence the reference to the raw object's equivalent.
    # @param [Workbook::Format] other_format
    # @return [Workbook::Format] a new resulting Workbook::Format
    def merge(other_format)
      self.remove_all_raws!
      self.merge_hash(other_format)
    end

    # Applies the formatting options of self with another, removes as a consequence the reference to the raw object's equivalent.
    # @param [Workbook::Format] other_format
    # @return [Workbook::Format] self
    def merge!(other_format)
      self.remove_all_raws!
      self.merge_hash!(other_format)
    end

    # returns an array of all formats this style is inheriting from (including itself)
    # @return [Array<Workbook::Format>] an array of Workbook::Formats
    def formats
      formats=[]
      f = self
      formats << f
      while f.parent
        formats << f.parent
        f = f.parent
      end
      formats.reverse
    end

    # returns an array of all format-names this style is inheriting from (and this style)
    # @return [Array<String>] an array of Workbook::Formats
    def all_names
      formats.collect{|a| a.name}
    end

    # Applies the formatting options of self with its parents until no parent can be found
    # @return [Workbook::Format] new Workbook::Format that is the result of merging current style with all its parent's styles.
    def flattened
      ff=Workbook::Format.new()
      formats.each{|a| ff.merge!(a) }
      return ff
    end

    # Formatting is sometimes the only way to detect the cells' type.
    def derived_type
      if self[:numberformat]
        if self[:numberformat].to_s.match("h")
          :time
        elsif self[:numberformat].to_s.match(/y/i)
          :date
        end
      end
    end
  end
end
