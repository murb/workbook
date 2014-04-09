# -*- encoding : utf-8 -*-
require 'rubyXL'
require 'workbook/readers/xls_shared'


module Workbook
  module Readers
    module XlsxReader
      include Workbook::Readers::XlsShared

      # Load method for .xlsm files, an office open file format, hence compatible with .xlsx (it emphasizes that it contains macros)
      #
      # @param [String, File] file_obj   a string with a reference to the file to be written to
      def load_xlsm file_obj
        self.load_xlsx file_obj
      end
      def load_xlsx file_obj
        file_obj = file_obj.path if file_obj.is_a? File
        sp = RubyXL::Parser.parse(file_obj)
        template.add_raw sp
        parse_xlsx sp
      end

      def parse_xlsx xlsx_spreadsheet=template.raws[RubyXL::Workbook], options={}
        options = {:additional_type_parsing=>false}.merge options
        #number_of_worksheets = xlsx_spreadsheet.worksheets.count
        xlsx_spreadsheet.worksheets.each_with_index do |worksheet, si|
          s = create_or_open_sheet_at(si)
          col_widths = []
          begin
            col_widths = xlsx_spreadsheet.worksheets.first.cols.collect{|a| a[:attributes][:width].to_f if a[:attributes]}
          rescue
            # Column widths couldn't be read, no big deal...
          end

          worksheet.each_with_index do |row, ri|
            if row
              r = s.table.create_or_open_row_at(ri)
              row.cells.each_with_index do |cell,ci|
                if cell.nil?
                  r[ci] = Workbook::Cell.new nil
                else
                  r[ci] = Workbook::Cell.new cell.value
                  r[ci].parse!
                  xls_format = cell.style_index
                  col_width = nil

                  if ri == 0
                    col_width = col_widths[ci]
                  end
                  f = template.create_or_find_format_by "style_index_#{cell.style_index}", col_width
                  f[:width]= col_width
                  background_color = cell.fill_color
                  background_color = (background_color.length == 8) ? background_color[2..8] : background_color #ignoring alpha for now.
                  f[:background_color] = "##{background_color}"

                  f[:number_format] = ms_formatting_to_strftime(cell.number_format)
                  # f[:font_family] = cell.font_name
                  # f[:color] = "##{cell.font_color}"

                  f.add_raw xls_format

                  r[ci].format = f
                end
              end
            end
          end
        end
      end
    end
  end
end
