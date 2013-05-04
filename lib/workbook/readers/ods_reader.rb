# -*- encoding : utf-8 -*-
module Workbook
  module Readers
    module OdsReader
      
      # reads self with and ods-type content.xml
      # @param [String,File] file_obj a file or file reference
      # @return [Workbook::Book] self
      def load_ods file_obj
        file_obj = file_obj.path if file_obj.is_a? File
        content = ""
        Zip::ZipFile.open(file_obj) do |zipfile|
          zipfile.entries.each do |file|
            content = zipfile.read(file.name) if file.name == "content.xml"  
          end
        end
        content = Nokogiri.XML(content)
        template.add_raw content
        parse_ods content
      end
      
      # updates self with and ods-type content.xml
      # @param [Nokogiri::XML::Document] ods_spreadsheet nokogirified content.xml
      # @return [Workbook::Book] self
      def parse_ods ods_spreadsheet=template.raws[Nokogiri::XML::Document], options={}
        options = {:additional_type_parsing=>false}.merge options
        ods_spreadsheet.xpath("//office:body/office:spreadsheet").each_with_index do |sheet,sheetindex|
          s = self.create_or_open_sheet_at(sheetindex)
          sheet.xpath("table:table").each_with_index do |table,tableindex|
            t = s.create_or_open_table_at(tableindex)
            table.xpath("table:table-row").each_with_index do |row,rowindex|
              row = row.xpath("table:table-cell").collect do |cell|
                c = Workbook::Cell.new()
                
                valuetype = cell.xpath('@office:value-type').to_s
                value = cell.xpath("text:p//text()").to_s
                value = value == "" ? nil : value
                case valuetype
                when 'float'
                  value = cell.xpath("@office:value").to_s.to_f
                when 'integer'
                  value = cell.xpath("@office:value").to_s.to_i
                when 'date'
                  value = DateTime.parse(cell.xpath("@office:date-value").to_s)
                end                
                c.value = value 
              end
              t << Workbook::Row.new(row)
            end
          end
        end
        return self
      end
    end
  end
end
