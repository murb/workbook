# -*- encoding : utf-8 -*-
require 'rubyXL'
require 'workbook/readers/xls_shared'

module Workbook
  module Writers
    module XlsxWriter

      # Generates an Spreadsheet (from the spreadsheet gem) in order to build an XlS
      #
      # @param [Hash] options A hash with options (unused so far)
      # @return [Spreadsheet] A Spreadsheet object, ready for writing or more lower level operations
      def to_xlsx options={}
        book = init_xlsx_spreadsheet_template
        book.theme = RubyXL::Theme.new unless book.theme #workaround bug in rubyxl
        book.worksheets.pop(book.worksheets.count - self.count) if book.worksheets and book.worksheets.count > self.count
        self.each_with_index do |s,si|
          xls_sheet = xlsx_sheet(si)
          xls_sheet.sheet_name = s.name

          s.table.each_with_index do |r, ri|
            r.each_with_index do |c, ci|
              xls_sheet.add_cell(ri, ci, c.value)

            end
          end
          (xls_sheet.count + 1 - s.table.count).times do |time|
            row_to_remove = s.table.count+time

            xls_sheet.delete_row(row_to_remove)
            # xls_sheet.row_updated(row_to_remove, xls_sheet.row(row_to_remove))
          end
          # xls_sheet.updated_from(s.table.count)
          # xls_sheet.dimensions

        end
        book
      end

      # # Generates an Spreadsheet (from the spreadsheet gem) in order to build an XlS
      # #
      # # @param [Workbook::Format, Hash] f A Workbook::Format or hash with format-options (:font_weight, :rotation, :background_color, :number_format, :text_direction, :color, :font_family)
      # # @return [Spreadsheet::Format] A Spreadsheet format-object, ready for writing or more lower level operations
      # def format_to_xlsx_format f
      #   xlsfmt = nil
      #   unless f.is_a? Workbook::Format
      #     f = Workbook::Format.new f
      #   end
      #   xlsfmt = f.return_raw_for Spreadsheet::Format
      #   unless xlsfmt
      #     xlsfmt=Spreadsheet::Format.new :weight=>f[:font_weight]
      #     xlsfmt.rotation = f[:rotation] if f[:rotation]
      #     xlsfmt.pattern_fg_color = html_color_to_xls_color(f[:background_color]) if html_color_to_xls_color(f[:background_color])
      #     xlsfmt.pattern = 1 if html_color_to_xls_color(f[:background_color])
      #     xlsfmt.number_format = strftime_to_ms_format(f[:number_format]) if f[:number_format]
      #     xlsfmt.text_direction = f[:text_direction] if f[:text_direction]
      #     xlsfmt.font.name = f[:font_family].split.first if f[:font_family]
      #     xlsfmt.font.family = parse_font_family(f) if f[:font_family]
      #     color = html_color_to_xls_color(f[:color])
      #     xlsfmt.font.color = color if color
      #     f.add_raw xlsfmt
      #   end
      #   return xlsfmt
      # end
      #
      # # Parses right font-family name
      # #
      # # @param [Workbook::Format, hash] format to parse
      # def parse_font_family(format)
      #   font = format[:font_family].to_s.split.last
      #   valid_values = [:none,:roman,:swiss,:modern,:script,:decorative]
      #   if valid_values.include?(font)
      #     return font
      #   elsif valid_values.include?(font.to_s.downcase.to_sym)
      #     return font.to_s.downcase.to_sym
      #   else
      #     font = font.to_s.downcase.strip
      #     translation = {
      #       "arial"=>:swiss,
      #       "times"=>:roman,
      #       "times new roman"=>:roman
      #     }
      #     tfont = translation[font]
      #     return tfont ? tfont : :none
      #   end
      # end
      #

      # Write the current workbook to Microsoft Excel format (using the spreadsheet gem)
      #
      # @param [String] filename
      # @param [Hash] options   see #to_xls
      def write_to_xlsx filename="#{title}.xlsx", options={}
        if to_xlsx(options).write(filename)
          return filename
        end
      end

      def xlsx_sheet a
        if xlsx_template.worksheets[a]
          return xlsx_template.worksheets[a]
        else
          xlsx_template.create_worksheet
          self.xls_sheet a
        end
      end

      def xlsx_template
        return template.raws[RubyXL::Workbook]
      end

      def init_xlsx_spreadsheet_template
        if self.xlsx_template.is_a? RubyXL::Workbook
          return self.xlsx_template
        else
          t = RubyXL::Workbook.new
          template.add_raw t
          return t
        end
      end
    end
  end
end
