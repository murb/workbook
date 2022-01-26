# frozen_string_literal: true
# frozen_string_literal: true

require "spreadsheet"

module Workbook
  module Writers
    module HtmlWriter
      # Generates an HTML table ()
      #
      # @param [Hash] options A hash with options
      # @return [String] A String containing the HTML code
      def to_html options = {}
        builder = Nokogiri::XML::Builder.new { |doc|
          doc.html {
            doc.body {
              each do |sheet|
                doc.h1 do
                  doc.text sheet.name
                end
                sheet.each do |table|
                  doc.h2 do
                    doc.text table.name
                  end
                  doc << table.to_html(options)
                end
              end
            }
          }
        }
        builder.doc.to_xhtml
      end

      # Write the current workbook to HTML format
      #
      # @param [String] filename
      # @param [Hash] options   see #to_xls
      # @return [String] filename

      def write_to_html filename = "#{title}.html", options = {}
        File.open(filename, "w") { |f| f.write(to_html(options)) }
        filename
      end
    end

    module HtmlTableWriter
      # Generates an HTML table ()
      #
      # @param [Hash] options A hash with options
      # @return [String] A String containing the HTML code, most importantly `:style_with_inline_css` (default false)
      def to_html options = {}
        options = {style_with_inline_css: false}.merge(options)
        builder = Nokogiri::XML::Builder.new { |doc|
          doc.table do
            doc.thead do
              if header
                doc.tr do
                  header.each do |cell|
                    th_options = build_cell_options cell, options.merge(classnames: [cell.to_sym], data: {key: cell.to_sym})
                    unless cell.value.instance_of?(Workbook::NilValue)
                      doc.th(th_options) do
                        doc.text cell.value
                      end
                    end
                  end
                end
              end
            end
            doc.tbody do
              each do |row|
                unless row.header?
                  doc.tr do
                    row.each do |cell|
                      td_options = build_cell_options cell, options
                      unless cell.value.instance_of?(Workbook::NilValue)
                        doc.td(td_options) do
                          doc.text cell.value
                        end
                      end
                    end
                  end
                end
              end
            end
          end
        }
        builder.doc.to_xhtml
      end

      def build_cell_options cell, options = {}
        classnames = cell.format.all_names
        classnames += options[:classnames] if options[:classnames]
        classnames = classnames.join(" ").strip
        td_options = classnames != "" ? {class: classnames} : {}
        options[:data]&.each do |key, value|
          td_options = td_options.merge({"data-#{key}".to_sym => value})
        end
        td_options = td_options.merge({style: cell.format.to_css}) if options[:style_with_inline_css] && (cell.format.to_css != "")
        td_options = td_options.merge({colspan: cell.colspan}) if cell.colspan
        td_options = td_options.merge({rowspan: cell.rowspan}) if cell.rowspan
        td_options
      end
    end
  end
end
