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
          workbook_sheet = self.create_or_open_sheet_at(sheetindex)
          sheet.xpath("table:table").each_with_index do |table,tableindex|
            parse_local_table(workbook_sheet,table,tableindex)
          end
        end
        return self
      end

      #parse the contents of an entire table by parsing every row in it and adding it to the table
      def parse_local_table(sheet,table,tableindex)
        puts table.to_xml
        local_table = sheet.create_or_open_table_at(tableindex)
        local_table.name = table.xpath("@table:name").to_s
        #column_count = get_column_count(table)
        table.xpath("table:table-row").each do |row|
          local_table << parse_local_row(row)
        end
      end

      #set column count
      def get_column_count(table)
        init_column_count = table.xpath("table:table-column").count
        first_row = table.xpath("table:table-row").first
        cells = first_row.xpath("table:table-cell|table:covered-table-cell")
        column_count = 0
        cells.each do |cell|
          if cell.xpath('@table:number-columns-spanned').children.size>0
            column_count +=cell.xpath('@table:number-columns-spanned').children[0].inner_text.to_i
          else
            column_count +=1
          end
        end
        column_count
      end

      #parse the contents of an entire row by parsing every cell in it and adding it to the row
      def parse_local_row(row)
        cells = row.xpath("table:table-cell|table:covered-table-cell")
        workbook_row = Workbook::Row.new
        cells.each do |cell|
          @cell = cell
          repeat = get_repeat
          workbook_cell = Workbook::Cell.new()
          workbook_cell.value = @cell.nil? ? nil : parse_local_cell(workbook_cell)
          repeat.times do
            workbook_row << workbook_cell
          end
        end
        return workbook_row
      end

      def get_repeat
        pre_set = @cell.xpath('@table:number-columns-repeated').to_s
        return 1 if (pre_set.nil? || pre_set=="") # if not present, don't repeat.
        return 1 unless "#{pre_set.to_i}"=="#{pre_set}" # return 1 if it's not a valid integer
        return 1 if pre_set.to_i < 1 # return 1, negative repeats make no sense
        return pre_set.to_i
      end

      #parse the contents of a single cell
      def parse_local_cell(workbook_cell)
        return Workbook::NilValue.new(:covered) if @cell.name == "covered-table-cell"
        set_cell_attributes(workbook_cell)
        valuetype = @cell.xpath('@office:value-type').to_s
        parse_local_value(valuetype)
      end

      # Sets cell attributes for rowspan, colspan and format
      def set_cell_attributes(workbook_cell)
        workbook_cell.format = self.template.formats[@cell.xpath('@table:style-name').to_s]
        workbook_cell.colspan= @cell.xpath('@table:number-columns-spanned').to_s
        workbook_cell.rowspan= @cell.xpath('@table:number-rows-spanned').to_s
      end

      # Sets value in right context type
      def parse_local_value(valuetype)
        value = CGI.unescapeHTML(@cell.xpath("text:p//text()").to_s)
        value = (value=="") ? nil : value
        case valuetype
        when 'float'
          value = @cell.xpath("@office:value").to_s.to_f
        when 'integer'
          value = @cell.xpath("@office:value").to_s.to_i
        when 'date'
          value = DateTime.parse(@cell.xpath("@office:date-value").to_s)
        end
        value
      end

    end
  end
end
