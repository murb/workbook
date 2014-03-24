# -*- encoding : utf-8 -*-
require 'spreadsheet'
require 'workbook/readers/xls_shared'


module Workbook
  module Readers
    module XlsReader
      include Workbook::Readers::XlsShared

      def load_xls file_obj
        begin
          sp = Spreadsheet.open(file_obj, 'rb')
          template.add_raw sp
          parse_xls sp
        rescue Ole::Storage::FormatError => e
          begin
            # Assuming it is a tab separated txt inside .xls
            open(file_obj.path, 'txt')
          rescue
            raise e
          end
        end

      end


      def parse_xls xls_spreadsheet=template.raws[Spreadsheet::Excel::Workbook], options={}
        options = {:additional_type_parsing=>true}.merge options
        number_of_worksheets = xls_spreadsheet.worksheets.count
        (0..number_of_worksheets-1).each do |si|
          xls_sheet = xls_spreadsheet.worksheets[si]
          begin
            number_of_rows = xls_spreadsheet.worksheets[si].count
            s = create_or_open_sheet_at(si)
            s.name = xls_sheet.name
            (0..number_of_rows-1).each do |ri|
              xls_row = xls_sheet.row(ri)
              r = s.table.create_or_open_row_at(ri)
              col_widths = xls_sheet.columns.collect{|c| c.width if c}
              xls_row.each_with_index do |xls_cell,ci|

                begin
                  r[ci] = Workbook::Cell.new xls_cell
                  r[ci].parse!
                rescue ArgumentError => e
                  if e.message.match('not a Spreadsheet::Formula')
                    v = xls_cell.value
                    if v.class == Float and xls_row.format(ci).date?
                      xls_row[ci] = v
                      v = xls_row.datetime(ci)
                    end
                    if v.is_a? Spreadsheet::Excel::Error
                      v = "----!"
                    end
                    r[ci] = Workbook::Cell.new v
                  elsif e.message.match('not a Spreadsheet::Link')
                    r[ci] = Workbook::Cell.new xls_cell.to_s
                  elsif e.message.match('not a Spreadsheet::Link')
                    r[ci] = Workbook::Cell.new xls_cell.to_s
                  elsif e.message.match('not a Spreadsheet::Excel::Error')
                    r[ci] = "._."
                  else
                    r[ci] = "._."  # raise e (we're going to be silent for now)
                  end
                end
                xls_format = xls_row.format(ci)
                col_width = nil

                if ri == 0
                  col_width = col_widths[ci]
                end
                f = template.create_or_find_format_by "object_id_#{xls_format.object_id}",col_width
                f[:width]= col_width
                f[:rotation] = xls_format.rotation if xls_format.rotation
                f[:background_color] = xls_color_to_html_hex(xls_format.pattern_fg_color)
                f[:number_format] = ms_formatting_to_strftime(xls_format.number_format)
                f[:text_direction] = xls_format.text_direction
                f[:font_family] = "#{xls_format.font.name}, #{xls_format.font.family}"
                f[:font_weight] = xls_format.font.weight
                f[:color] = xls_color_to_html_hex(xls_format.font.color)

                f.add_raw xls_format

                r[ci].format = f
              end
            end
          rescue TypeError
            puts "WARNING: Failed at worksheet (#{si})... ignored"
            #ignore SpreadsheetGem can be buggy...
          end
        end
      end

      private
      def xls_color_to_html_hex color_sym
        Workbook::Book::XLS_COLORS[color_sym] ? Workbook::Book::XLS_COLORS[color_sym] : "#000000"
      end
    end
  end
end
