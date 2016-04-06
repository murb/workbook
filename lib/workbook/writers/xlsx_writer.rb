# -*- encoding : utf-8 -*-
require 'axlsx'
require 'workbook/readers/xls_shared'

module Workbook
  module Writers
    module XlsxWriter

      # Generates an axlsx doc, ready for export to XLSX
      #
      # @param [Hash] options A hash with options (unused so far)
      # @return [Axlsx::Package] An object, ready for writing or more lower level operations
      def to_xlsx options={}
        formats_to_xlsx_format
        book = init_xlsx_spreadsheet_template.workbook
        book.worksheets.pop(book.worksheets.count - self.count) if book.worksheets and book.worksheets.count > self.count
        self.each_with_index do |s,si|
          xlsx_sheet = xlsx_sheet(si)
          xlsx_sheet.name = s.name
          s.table.each_with_index do |r, ri|
            xlsx_row = xlsx_sheet[ri] ? xlsx_sheet[ri] : xlsx_sheet.add_row
            xlsx_row.height = 16
            xlsx_row_a = xlsx_row.to_ary
            r.each_with_index do |c, ci|
              xlsx_row.add_cell(c.value) unless xlsx_row_a[ci]
              xlsx_cell = xlsx_row_a[ci]
              xlsx_cell.value = c.value
              if c.format?
                format_to_xlsx_format(c.format) unless c.format.raws[Fixnum]
                xlsx_cell.style = c.format.raws[Fixnum]
              end
            end
            xlsx_sheet.send(:update_column_info, xlsx_row.cells, [])
          end
          (xlsx_sheet.rows.count - s.table.count).times do |time|
            xlsx_sheet.rows.pop
          end
        end
        init_xlsx_spreadsheet_template
      end

      # Generates an string ready to be streamed as XLSX
      #
      # @param [Hash] options A hash with options (unused so far)
      # @return [String] A string, ready for streaming, e.g. `send_data workbook.stream_xlsx`
      def stream_xlsx options = {}
        to_xlsx(options).to_stream.string
      end

      # Write the current workbook to Microsoft Excel's XML format (using the Axlsx gem)
      #
      # @param [String] filename
      # @param [Hash] options   see #to_xlsx
      def write_to_xlsx filename="#{title}.xlsx", options={}
        if to_xlsx(options).serialize(filename)
          return filename
        end
      end

      def xlsx_sheet a
        if xlsx_template.workbook.worksheets[a]
          return xlsx_template.workbook.worksheets[a]
        else
          xlsx_template.workbook.add_worksheet
          self.xlsx_sheet a
        end
      end

      def xlsx_template
        return template.raws[Axlsx::Package]
      end

      def init_xlsx_spreadsheet_template
        if self.xlsx_template.is_a? Axlsx::Package
          return self.xlsx_template
        else
          t = Axlsx::Package.new
          template.add_raw t
          template.set_default_formats!
          return t
        end
      end

      def formats_to_xlsx_format
        template.formats.each do |n,v|
          v.each do | t,s |
            format_to_xlsx_format(s)
          end
        end
      end

      def format_to_xlsx_format f
        xlsfmt = nil
        unless f.is_a? Workbook::Format
          f = Workbook::Format.new f
        end

        xlsfmt={}
        xlsfmt[:fg_color] = "FF#{f[:color].to_s.upcase}".gsub("#",'') if f[:color]
        xlsfmt[:b] = true if f[:font_weight].to_s == "bold" or f[:font_weight].to_i >= 600 or f[:"font_style"].to_s.match "oblique"
        xlsfmt[:i] = true if f[:font_style].to_s == "italic"
        xlsfmt[:u] = true if f[:text_decoration].to_s.match("underline")
        xlsfmt[:bg_color] = f[:background_color] if f[:background_color]
        xlsfmt[:format_code] = strftime_to_ms_format(f[:number_format]) if f[:number_format]
        xlsfmt[:font_name] = f[:font_family].split.first if f[:font_family]
        xlsfmt[:family] = parse_font_family(f) if f[:font_family]
        f.add_raw init_xlsx_spreadsheet_template.workbook.styles.add_style(xlsfmt)
        # wb.styles{|a| p a.add_style({}).class }

        f.add_raw xlsfmt
        return xlsfmt
      end
    end
  end
end
