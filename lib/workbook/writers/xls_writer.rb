# -*- encoding : utf-8 -*-
# frozen_string_literal: true
require 'spreadsheet'

module Workbook
  module Writers
    module XlsWriter

      # Generates an Spreadsheet (from the spreadsheet gem) in order to build an xls
      #
      # @param [Hash] options A hash with options (unused so far)
      # @return [Spreadsheet] A Spreadsheet object, ready for writing or more lower level operations
      def to_xls options={}
        book = init_spreadsheet_template
        self.each_with_index do |s,si|
          xls_sheet = xls_sheet(si)
          xls_sheet.name = s.name

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
          (xls_sheet.last_row_index + 1 - s.table.count).times do |time|
            row_to_remove = s.table.count+time
            remove_row(xls_sheet,row_to_remove)
          end
          xls_sheet.updated_from(s.table.count)
          xls_sheet.dimensions

        end
        # kind of a hack, deleting by popping from xls worksheet results in errors in MS Excel (not LibreOffice)
        # book.worksheets.pop(book.worksheets.count - self.count) if book.worksheets and book.worksheets.count > self.count
        book.worksheets.each_with_index do |xls_sheet, si|
          if self[si]
            xls_sheet.visibility = :visible
          else
            xls_sheet.visibility = :strong_hidden
            #also make sure all data is removed, in case someone finds out about this 'trick'
            xls_sheet.name = "RemovedSheet#{si}"
            (xls_sheet.last_row_index + 1).times do |row_index|
              remove_row(xls_sheet,row_index)
            end
          end
        end
        # even after removing the worksheet's content... which solved some incompatibilities, but not for popping worksheets
        # book.worksheets.pop(book.worksheets.count - self.count) if book.worksheets and book.worksheets.count > self.count
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
          xlsfmt.font.family = parse_font_family(f) if f[:font_family]
          color = html_color_to_xls_color(f[:color])
          xlsfmt.font.color = color if color
          f.add_raw xlsfmt
        end
        return xlsfmt
      end

      # Parses right font-family name
      #
      # @param [Workbook::Format, hash] format to parse
      def parse_font_family(format)
        font = format[:font_family].to_s.split.last
        valid_values = [:none,:roman,:swiss,:modern,:script,:decorative]
        if valid_values.include?(font)
          return font
        elsif valid_values.include?(font.to_s.downcase.to_sym)
          return font.to_s.downcase.to_sym
        else
          font = font.to_s.downcase.strip
          translation = {
            "arial"=>:swiss,
            "times"=>:roman,
            "times new roman"=>:roman
          }
          tfont = translation[font]
          return tfont ? tfont : :none
        end
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

      private

      def remove_row(xls_sheet,row_index)
        xls_sheet.row(row_index).each_with_index do |c, ci|
          xls_sheet.row(row_index)[ci]=nil
        end
        xls_sheet.delete_row(row_index)
        xls_sheet.row_updated(row_index, xls_sheet.row(row_index))
      end
    end
  end
end
