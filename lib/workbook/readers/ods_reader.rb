# -*- encoding : utf-8 -*-

# TODO: rods (require 'rods') works weird... going to write my own parser.

module Workbook
  module Readers
    module OdsReader
      def load_ods file_obj
        file_obj = file_obj.path if file_obj.is_a? File
        sp = Rods.new(file_obj)
        template.add_raw sp
        parse_ods sp
      end
      
      def parse_ods_cell cell
        value, type = cell.read_cell
        if type == "string"
          return value
        elsif type == "float"
          if value.to_f - value.to_i == 0.0
            value = value.to_i
          else
            value = value.to_f
          end
        elsif type == "date"
          year,month,day= cell.attributes["office:date-value"].split('-')
          value = Time.new(year.to_i,month.to_i,day.to_i)
        end
        return Workbook::Cell.new(value)
      end
      
      def parse_ods_row sheet, row_index
        values = Workbook::Row.new
        cell=sheet.getCell(row_index,1)
        values.push parse_ods_cell(cell)
        while(cell=mySheet.getNextExistentCell(cell)) do
          values.push parse_ods_cell(cell)
        end
        return values
      end
      
      def parse_ods ods_spreadsheet=template.raws[Rods], options={}
        options = {:additional_type_parsing=>false}.merge options
        row_index = 1
        row=ods_spreadsheet.getRow(1)
        wb = Workbook::Book.new
        table = wb.sheet.first
        table.push(parse_ods_row(ods_spreadsheet,row))
        while(row=mySheet.getNextExistentRow(row))
          row_index += 1
          table.push(parse_ods_row(ods_spreadsheet,row_index))
        end
        return wb
      end
    end
  end
end
