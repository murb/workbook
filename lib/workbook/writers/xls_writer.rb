require 'spreadsheet'

module Workbook
  module Writers
    module XlsWriter

      # Generates an Spreadsheet (from the spreadsheet gem) in order to build an XlS
      # 
      # @param [Hash] options A hash with options (unused so far)
      # @return [Spreadsheet] A Spreadsheet object, ready for writing or more lower level operations
      def to_xls options={}
        book = init_spreadsheet_template
        self.each_with_index do |s,si|
          xls_sheet = book.worksheet si
          xls_sheet = book.create_worksheet if xls_sheet == nil
          s.table.each_with_index do |r, ri|
            xls_sheet.row(ri).height= r.format[:height] if r.format
            r.each_with_index do |c, ci|
              if c
                if r.first?
                  xls_sheet.columns[ci] ||= Spreadsheet::Column.new(ci,nil)
                  xls_sheet.columns[ci].width= c.format[:width]
                end
                xls_sheet.row(ri)[ci] = c.value
                xls_sheet.row(ri).set_format(ci, format_to_xls_format(c.format))
              end
            end
          end
        end
        book
      end
      
      # Generates an Spreadsheet (from the spreadsheet gem) in order to build an XlS
      # 
      # @param [Workbook::Format, Hash] f A Workbook::Format or hash with format-options (:font_weight, :rotation, :background_color, :number_format, :text_direction, :color, :font_family)
      # @return [Spreadsheet::Format] A Spreadsheet format-object, ready for writing or more lower level operations
      def format_to_xls_format f
        xlsfmt = nil
        unless f.is_a? Workbook::Format
          f = Workbook::Format.new f
        end
        xlsfmt = f.return_raw_for Spreadsheet::Format
        unless xlsfmt
          xlsfmt=Spreadsheet::Format.new :weight=>f[:font_weight]
          xlsfmt.rotation = f[:rotation] if f[:rotation] 
          xlsfmt.pattern_fg_color = html_color_to_xls_color(f[:background_color]) if html_color_to_xls_color(f[:background_color])
          xlsfmt.pattern = 1 if html_color_to_xls_color(f[:background_color])
          xlsfmt.number_format = strftime_to_ms_format(f[:number_format]) if f[:number_format]
          xlsfmt.text_direction = f[:text_direction] if f[:text_direction]
          xlsfmt.font.name = f[:font_family].split.first if f[:font_family]
          xlsfmt.font.family = f[:font_family].split.last if f[:font_family]
          xlsfmt.font.color = html_color_to_xls_color(f[:color]) if f[:color]
          f.add_raw xlsfmt
        end
        return xlsfmt
      end
      
      # Attempt to convert html-hex color value to xls color number
      #
      # @param [String] hex color
      # @return [String] xls color
      def html_color_to_xls_color hex
        Workbook::Readers::XlsShared::XLS_COLORS.each do |k,v|
          return k if (v == hex or (hex and hex != "" and k == hex.to_sym))
        end
        return nil
      end
      
      # Converts standard (ruby/C++/unix/...) strftime formatting to MS's formatting
      # 
      # @param [String, nil] numberformat (nil returns nil)
      # @return [String, nil]
      def strftime_to_ms_format numberformat
        return nil if numberformat.nil?
        return numberformat.gsub('%Y','yyyy').gsub('%A','dddd').gsub('%B','mmmm').gsub('%a','ddd').gsub('%b','mmm').gsub('%y','yy').gsub('%d','dd').gsub('%m','mm').gsub('%y','y').gsub('%y','%%y').gsub('%e','d')
      end
      
      # Write the current workbook to Microsoft Excel format (using the spreadsheet gem)
      #
      # @param [String] filename
      # @param [Hash] options   see #to_xls 
      def write_to_xls filename="#{title}.xls", options={}
        if to_xls(options).write(filename)
          return filename
        end
      end
    
      def xls_sheet a
        if xls_template.worksheet(a)
          return xls_template.worksheet(a)
        else
          xls_template.create_worksheet
          self.xls_sheet a
        end
      end
      
      def xls_template
        return template.raws[Spreadsheet::Excel::Workbook] ? template.raws[Spreadsheet::Excel::Workbook] : template.raws[Spreadsheet::Workbook]
      end
            
      def init_spreadsheet_template
        if self.xls_template.is_a? Spreadsheet::Workbook
          return self.xls_template
        else
          t = Spreadsheet::Workbook.new
          template.add_raw t
          return t
        end
      end
    end
  end
end