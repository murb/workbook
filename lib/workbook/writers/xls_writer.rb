require 'spreadsheet'

module Workbook
  module Writers
    module XlsWriter
      
      def to_xls options={}
        options = {:rewrite_header=>false}.merge options
        book = init_spreadsheet_template
        self.each_with_index do |s,si|
          xls_sheet = book.worksheet si
          xls_sheet = book.create_worksheet if xls_sheet == nil
          s.table.each_with_index do |r, ri|
            write_row = false
            if r.header?
              if options[:rewrite_header] == true
                write_row = true
              end
            else
              write_row = true
            end
            if write_row
              xls_sheet.row(ri).replace r.collect{|c| c.value }
            end
          end
        end
        book
      end
      
      def write_to_xls options={}
        filename = options[:filename] ? options[:filename] : "#{title}.xls"
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