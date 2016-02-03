# -*- encoding : utf-8 -*-
require 'roo'
require 'workbook/readers/xls_shared'


module Workbook
  module Readers
    module XlsxReader
      include Workbook::Readers::XlsShared

      # Load method for .xlsm files, an office open file format, hence compatible with .xlsx (it emphasizes that it contains macros)
      #
      # @param [String, File] file_obj   a string with a reference to the file to be written to
      def load_xlsm file_obj, options={}
        self.load_xlsx file_obj, options
      end
      def load_xlsx file_obj, options={}
        file_obj = file_obj.path if file_obj.is_a? File
        # file_obj = file_obj.match(/^\/(.*)/) ? file_obj : "./#{file_obj}"
        # p "opening #{file_obj}"
        sp = Roo::Excelx.new(file_obj)
        template.add_raw sp, raw_object_class: Roo::Spreadsheet
        parse_xlsx sp
      end

      def parse_xlsx xlsx_spreadsheet=template.raws[Roo::Spreadsheet], options={}
        options = {:additional_type_parsing=>false}.merge options
        sheet_index = 0
        xlsx_spreadsheet.each_with_pagename do |sheet_name, sheet|
          s = create_or_open_sheet_at(sheet_index)
          sheet.each_with_index do |row, rowi|
            s.table << row
          end
          sheet_index += 1
        end
      end
    end
  end
end
