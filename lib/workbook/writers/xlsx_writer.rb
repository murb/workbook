# frozen_string_literal: true

require "axlsx"
require "workbook/readers/xls_shared"

module Workbook
  module Writers
    module XlsxWriter
      CELL_TYPE_MAPPING = {
        decimal: :integer,
        integer: :integer,
        float: :float,
        string: :text,
        time: :time,
        date: :date,
        datetime: :time,
        boolean: :boolean,
        nil: :string
      }
      # Generates an axlsx doc, ready for export to XLSX
      #
      # @param [Hash] options A hash with options (unused so far)
      # @return [Axlsx::Package] An object, ready for writing or more lower level operations
      def to_xlsx options = {}
        formats_to_xlsx_format

        book = init_xlsx_spreadsheet_template.workbook
        book.worksheets.pop(book.worksheets.count - count) if book.worksheets && (book.worksheets.count > count)
        each_with_index do |s, si|
          xlsx_sheet = xlsx_sheet(si)
          xlsx_sheet.name = s.name
          s.table.each_with_index do |r, ri|
            xlsx_row = xlsx_sheet[ri] || xlsx_sheet.add_row
            xlsx_row.height = 16
            xlsx_row_a = xlsx_row.to_ary
            r.each_with_index do |c, ci|
              xlsx_row.add_cell(c.value) unless xlsx_row_a[ci]
              xlsx_cell = xlsx_row_a[ci]

              xlsx_cell.type = CELL_TYPE_MAPPING[c.cell_type]
              xlsx_cell.value = c.value

              if c.format? && c.format.raws[Integer]
                xlsx_cell.style =  c.format.raws[Integer]
              elsif c.format? && !(c.value.is_a?(Date) || !c.value.is_a?(DateTime) || !c.value.is_a?(Time))
                # TODO: disable formatting
                # xlsx_cell.style = format_to_xlsx_format(c.format)
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
      def write_to_xlsx filename = "#{title}.xlsx", options = {}
        if to_xlsx(options).serialize(filename)
          filename
        end
      end

      def xlsx_sheet a
        if xlsx_template.workbook.worksheets[a]
          xlsx_template.workbook.worksheets[a]
        else
          xlsx_template.workbook.add_worksheet
          xlsx_sheet a
        end
      end

      def xlsx_template
        template.raws[Axlsx::Package]
      end

      def init_xlsx_spreadsheet_template
        if xlsx_template.is_a? Axlsx::Package
          xlsx_template
        else
          t = Axlsx::Package.new
          template.add_raw t
          # template.set_default_formats!
          t
        end
      end

      def formats_to_xlsx_format
        template.formats.each do |n, v|
          v.each do |t, s|
            format_to_xlsx_format(s)
          end
        end
      end

      def make_sure_f_is_a_workbook_format f
        f.is_a?(Workbook::Format) ? f : Workbook::Format.new({}, f)
      end

      def format_to_xlsx_format f
        f = make_sure_f_is_a_workbook_format f

        xlsfmt = {}
        xlsfmt[:fg_color] = "FF#{f[:color].to_s.upcase}".delete("#") if f[:color]
        xlsfmt[:b] = true if (f[:font_weight].to_s == "bold") || (f[:font_weight].to_i >= 600) || f[:font_style].to_s.match("oblique")
        xlsfmt[:i] = true if f[:font_style].to_s == "italic"
        xlsfmt[:u] = true if f[:text_decoration].to_s.include?("underline")
        xlsfmt[:bg_color] = f[:background_color] if f[:background_color]
        xlsfmt[:format_code] = strftime_to_ms_format(f[:number_format]) if f[:number_format]
        xlsfmt[:font_name] = f[:font_family].split.first if f[:font_family]
        # xlsfmt[:family] = parse_font_family(f) if f[:font_family]

        style_reference = init_xlsx_spreadsheet_template.workbook.styles.add_style(xlsfmt)

        f.add_raw style_reference
        f.add_raw xlsfmt

        style_reference
      end
    end
  end
end
