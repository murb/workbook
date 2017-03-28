# -*- encoding : utf-8 -*-
require 'workbook/readers/xls_shared'

module Workbook
  module Readers
    module XlsxReader
      include Workbook::Readers::XlsShared

      def load_xlsm file_obj, options={}
        self.load_xlsx file_obj, options
      end
      def load_xlsx file_obj, options={}
        file_obj = file_obj.path if file_obj.is_a? File
        sheets = {}
        shared_string_file = ""
        styles = ""
        workbook = ""
        workbook_rels = ""
        Zip::File.open(file_obj) do |zipfile|
          zipfile.entries.each do |file|
            if file.name.match(/^xl\/worksheets\/(.*)\.xml$/)
              sheets[file.name.sub(/^xl\//,'')] = zipfile.read(file.name)
            elsif file.name.match(/xl\/sharedStrings.xml$/)
              shared_string_file = zipfile.read(file.name)
            elsif file.name.match(/xl\/workbook.xml$/)
              workbook = zipfile.read(file.name)
            elsif file.name.match(/xl\/_rels\/workbook.xml.rels$/)
              workbook_rels = zipfile.read(file.name)
            elsif file.name.match(/xl\/styles.xml$/)
              styles = zipfile.read(file.name)
            end
            # content = zipfile.read(file.name) if file.name == "content.xml"
          end
        end

        parse_xlsx_styles(styles)


        relation_file = {}
        Nokogiri::XML(workbook_rels).css("Relationships Relationship").each do |relship|
          relation_file[relship.attr("Id")]=relship.attr("Target")
        end

        @shared_strings = parse_shared_string_file(shared_string_file)

        Nokogiri::XML(workbook).css("sheets sheet").each do |sheet|
          name = sheet.attr("name")
          filename = relation_file[sheet.attr("r:id")]
          state = sheet.attr("state")
          if state != "hidden"
            sheet = sheets[filename]
            self.push parse_xlsx_sheet(sheet)
            self.last.name = name
          end
        end

        @shared_strings = nil
        self.each do |sheet|
          sheet.each do |table|
            table.trim!
          end
        end
      end

      def parse_xlsx_styles(styles)
        styles = Nokogiri::XML(styles)

        fonts = parse_xlsx_fonts styles
        backgrounds = extract_xlsx_backgrounds styles
        customNumberFormats = extract_xlsx_number_formats styles


        styles.css("cellXfs xf").each do |cellformat|
          hash = {}
          # <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">
          background_hash = backgrounds[cellformat.attr("applyFill").to_i]
          hash.merge!(background_hash) if background_hash

          font_hash = fonts[cellformat.attr("applyFill").to_i]
          hash.merge!(font_hash) if font_hash

          id = (cellformat.attr("numFmtId")).to_i
          if id >= 164
            hash[:numberformat] = customNumberFormats[id]
          else
            hash[:numberformat] = ms_formatting_to_strftime(id)
          end
          self.template.add_format Workbook::Format.new(hash)
        end
      end

      # Extracts fonts descriptors from styles.xml
      def parse_xlsx_fonts styles
        styles.css("fonts font").collect do |font|
          hash = {}
          hash[:font_family] = font.css("name").first.attr("val") if font.css("name")
          hash[:font_size] = font.css("sz").first.attr("val").to_i if font.css("name")
          hash
        end
      end

      # Extracts number formats from styles.xml
      def extract_xlsx_number_formats styles
        hash = {}
        styles.css("numFmts numFmt").each do |fmt|
          format_id = fmt.attr("numFmtId").to_i
          parsed_format_string = ms_formatting_to_strftime(fmt.attr("formatCode"))
          hash[format_id] = parsed_format_string
        end
        hash
      end

      def extract_xlsx_backgrounds styles
        styles.css("fills fill").collect do |fill|
          hash = {}
          patternFill = fill.css("patternFill").first
          # TODO: convert to html-hex
          hash[:background] = patternFill.attr("patternType") if patternFill
          hash
        end
      end

      def parse_shared_string_file file
        Nokogiri::XML(file).css("sst si t").collect{|a| a.text}
      end
      def parse_xlsx_sheet sheet_xml
        sheet = Workbook::Sheet.new
        table = sheet.table

        noko_xml = Nokogiri::XML(sheet_xml)

        rows = noko_xml.css('sheetData row').collect{|row| parse_xlsx_row(row)}
        rows.each do |row|
          table << row
        end

        columns = noko_xml.css('cols col').collect{|col| parse_xlsx_column(col) }
        table.columns = columns

        sheet
      end

      def parse_xlsx_column column
        col = Workbook::Column.new()
        col.width = column.attr("width").to_f
        col
      end
      def parse_xlsx_row row
        cells_with_pos = row.css('c').collect{|a| parse_xlsx_cell(a)}
        row = Workbook::Row.new
        cells_with_pos.each do |cell_with_pos|
          position = cell_with_pos[:position]
          col = position.match(/^[A-Z]*/).to_s
          row[col] = cell_with_pos[:cell]
        end
        row = pad_xlsx_row(row)
        row
      end

      def pad_xlsx_row(row)
        row.each_with_index do |cell, index|
          row[index]=Workbook::Cell.new(nil) if cell.nil? and !cell.is_a?(Workbook::Cell)
        end
        row
      end

      def parse_xlsx_cell cell
        # style_id = cell.attr('s')
        # p cell
        type = cell.attr('t')
        formatIndex = cell.attr('s').to_i
        position = cell.attr('r')
        formula = cell.css('f').text()
        value = cell.text
        fmt = template.formats[formatIndex]

        # puts type
        if type == "n" or type == nil
          if fmt.derived_type == :date
            value = xls_number_to_date(value)
          elsif fmt.derived_type == :time
            value = xls_number_to_time(value)
          elsif formula == "TRUE()"
            value = true
          elsif formula == "FALSE()"
            value = false
          elsif type == "n"
            value = value.match(/\./) ? value.to_f : value.to_i
          end
        elsif type == "s"
          value = @shared_strings[value.to_i]
        end
        cell = Workbook::Cell.new(value, format: fmt)
        cell.formula = formula
        {cell: cell, position: position}
      end
      def parse_xlsx
      end
    end
  end
end