require 'spreadsheet'

module Workbook
  module Writers
    module XlsWriter
      # TODO: find better way to dupe
      XLS_COLORS = {:xls_color_1=>'#000000',
                    :xls_color_2=>'#FFFFFF',
                    :xls_color_3=>'#FF0000',
                    :xls_color_4=>'#00FF00',
                    :xls_color_5=>'#0000FF',
                    :xls_color_6=>'#FFFF00',
                    :xls_color_7=>'#FF00FF',
                    :xls_color_8=>'#00FFFF',
                    :xls_color_9=>'#800000',
                    :xls_color_10=>'#008000',
                    :xls_color_11=>'#000080',
                    :xls_color_12=>'#808000',
                    :xls_color_13=>'#800080',
                    :xls_color_14=>'#008080',
                    :xls_color_15=>'#C0C0C0',
                    :xls_color_16=>'#808080',
                    :xls_color_17=>'#9999FF',
                    :xls_color_18=>'#993366',
                    :xls_color_19=>'#FFFFCC',
                    :xls_color_20=>'#CCFFFF',
                    :xls_color_21=>'#660066',
                    :xls_color_22=>'#FF8080',
                    :xls_color_23=>'#0066CC',
                    :xls_color_24=>'#CCCCFF',
                    :xls_color_25=>'#000080',
                    :xls_color_26=>'#FF00FF',
                    :xls_color_27=>'#FFFF00',
                    :xls_color_28=>'#00FFFF',
                    :xls_color_29=>'#800080',
                    :xls_color_30=>'#800000',
                    :xls_color_31=>'#008080',
                    :xls_color_32=>'#0000FF',
                    :xls_color_33=>'#00CCFF',
                    :xls_color_34=>'#CCFFFF',
                    :xls_color_35=>'#CCFFCC',
                    :xls_color_36=>'#FFFF99',
                    :xls_color_37=>'#99CCFF',
                    :xls_color_38=>'#FF99CC',
                    :xls_color_39=>'#CC99FF',
                    :xls_color_40=>'#FFCC99',
                    :xls_color_41=>'#3366FF',
                    :xls_color_42=>'#33CCCC',
                    :xls_color_43=>'#99CC00',
                    :xls_color_44=>'#FFCC00',
                    :xls_color_45=>'#FF9900',
                    :xls_color_46=>'#FF6600',
                    :xls_color_47=>'#666699',
                    :xls_color_48=>'#969696',
                    :xls_color_49=>'#003366',
                    :xls_color_50=>'#339966',
                    :xls_color_51=>'#003300',
                    :xls_color_52=>'#333300',
                    :xls_color_53=>'#993300',
                    :xls_color_54=>'#993366',
                    :xls_color_55=>'#333399',
                    :xls_color_56=>'#333333',
                    :black=>'#000000',
                    :white=>'#FFFFFF',
                    :red=>'#FF0000',
                    :green=>'#00FF00',
                    :blue=>'#0000FF',
                    :yellow=>'#FFFF00',
                    :magenta=>'#FF00FF',
                    :cyan=>'#00FFFF',
                    :border=>'#FFFFFF',
                    :text=>'#000000',
                    :lime=>'#00f94c'
      }
      
      # Generates an Spreadsheet (from the spreadsheet gem) in order to build an XlS
      # 
      # @params [Hash] A hash with options (unused so far)
      # @returns [Spreadsheet] A Spreadsheet object, ready for writing or more lower level operations
      def to_xls options={}
        book = init_spreadsheet_template
        self.each_with_index do |s,si|
          xls_sheet = book.worksheet si
          xls_sheet = book.create_worksheet if xls_sheet == nil
          s.table.each_with_index do |r, ri|
            xls_sheet.row(ri).height= r.format[:height] if r.format
            r.each_with_index do |c, ci|
              if c
                if r.header?
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
      
      def html_color_to_xls_color hex
        XLS_COLORS.each do |k,v|
          return k if (v == hex or (hex and hex != "" and k == hex.to_sym))
        end
        return nil
      end
      
      def strftime_to_ms_format numberformat
        return nil if numberformat.nil?
        numberformat.gsub('%Y','yyyy').gsub('%A','dddd').gsub('%B','mmmm').gsub('%a','ddd').gsub('%b','mmm').gsub('%y','yy').gsub('%d','dd').gsub('%m','mm').gsub('%y','y').gsub('%y','%%y').gsub('%e','d')
      end
      
      # Write the current workbook to Microsoft Excel format (using the spreadsheet gem)
      #
      # @param [String] the filename
      # @param [Hash] options, see #to_xls 
      
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