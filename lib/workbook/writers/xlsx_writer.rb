# -*- encoding : utf-8 -*-
require 'rubyXL'
require 'workbook/readers/xls_shared'

module Workbook
  module Writers
    module XlsxWriter

      # Generates an RubyXL doc, ready for export to XLSX
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

      # Generates an RubyXL doc, ready for export to XLSX
      #
      # @param [Hash] options A hash with options (unused so far)
      # @return [String] A string, ready for streaming, e.g. `send_data workbook.stream_xlsx`
      def stream_xlsx options = {}
        to_xlsx(options).stream.string
      end

      # Write the current workbook to Microsoft Excel's new format (using the RubyXL gem)
      #
      # @param [String] filename
      # @param [Hash] options   see #to_xlsx
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
