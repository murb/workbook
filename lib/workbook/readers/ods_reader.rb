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
        styles = ""
        Zip::ZipFile.open(file_obj) do |zipfile|
          zipfile.entries.each do |file|
            styles = zipfile.read(file.name) if file.name == "styles.xml"  
            
            content = zipfile.read(file.name) if file.name == "content.xml"  
          end
        end
        content = Nokogiri.XML(content)
        styles = Nokogiri.XML(styles)
        template.add_raw content
        parse_ods_style styles
        parse_ods content
        return self
      end
      
      def set_format_property format, property, value
        value.strip!
        format[property] = value if value and value != ""
      end
      
      def parse_ods_style parse_ods_style
        parse_ods_style.xpath("//style:style").each do |style|
          style_family = style.xpath("@style:family").to_s
          if style_family == "table-cell"
            format = Workbook::Format.new
            format.name = style.xpath("@style:name").to_s
            format.parent = self.template.formats[style.xpath("@style:parent-style-name").to_s]
            set_format_property format, :border, style.xpath("style:table-cell-properties/@fo:border").to_s
            set_format_property format, :vertical_align, style.xpath("style:table-cell-properties/@style:vertical-align").to_s.gsub("automatic","auto")
            set_format_property format, :padding, style.xpath("style:table-cell-properties/@fo:padding").to_s.gsub("automatic","auto")
            set_format_property format, :font, style.xpath("style:text-properties/style:font-name").to_s + " " + style.xpath("style:text-properties/fo:font-size").to_s + " " + style.xpath("style:text-properties/fo:font-weight").to_s
            set_format_property format, :color, style.xpath("style:text-properties/@fo:color").to_s
            set_format_property format, :background_color, style.xpath("style:table-cell-properties/@fo:background-color").to_s
            self.template.add_format(format)
          end
        end
      end
      
      # updates self with and ods-type content.xml
      # @param [Nokogiri::XML::Document] ods_spreadsheet nokogirified content.xml
      # @return [Workbook::Book] self
      def parse_ods ods_spreadsheet=template.raws[Nokogiri::XML::Document], options={}
        require 'cgi'
        
        options = {:additional_type_parsing=>false}.merge options
        # styles
        #puts ods_spreadsheet
        parse_ods_style ods_spreadsheet
        
        # data
        ods_spreadsheet.xpath("//office:body/office:spreadsheet").each_with_index do |sheet,sheetindex|
          s = self.create_or_open_sheet_at(sheetindex)
          sheet.xpath("table:table").each_with_index do |table,tableindex|
            t = s.create_or_open_table_at(tableindex)
            table.xpath("table:table-row").each_with_index do |row,rowindex|
              row = row.xpath("table:table-cell").collect do |cell|
                c = Workbook::Cell.new()
                cell_style_name = cell.xpath('@table:style-name').to_s
                c.format = self.template.formats[cell_style_name]
                valuetype = cell.xpath('@office:value-type').to_s
                value = CGI.unescapeHTML(cell.xpath("text:p//text()").to_s)
                value = (value == "") ? nil : value
                case valuetype
                when 'float'
                  value = cell.xpath("@office:value").to_s.to_f
                when 'integer'
                  value = cell.xpath("@office:value").to_s.to_i
                when 'date'
                  value = DateTime.parse(cell.xpath("@office:date-value").to_s)
                end           
                c.value = value
                c
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
