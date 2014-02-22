# -*- encoding : utf-8 -*-
require 'spreadsheet'

module Workbook
  module Writers
    module HtmlWriter

      # Generates an HTML table ()
      #
      # @param [Hash] options A hash with options
      # @return [String] A String containing the HTML code
      def to_html options={}
        options = {:style_with_inline_css=>false}.merge(options)
        builder = Nokogiri::XML::Builder.new do |doc|
          doc.html {
            doc.body {
              self.each{|sheet|
                doc.h1 {
                  doc.text sheet.name
                }
                sheet.each{|table|
                  doc.h2 {
                    doc.text table.name
                  }
                  doc << table.to_html(options)
                }
              }
            }
          }
        end
        return builder.doc.to_xhtml
      end

      # Write the current workbook to HTML format
      #
      # @param [String] filename
      # @param [Hash] options   see #to_xls
      # @return [String] filename

      def write_to_html filename="#{title}.html", options={}
        File.open(filename, 'w') {|f| f.write(to_html(options)) }
        return filename
      end
    end

    module HtmlTableWriter
      # Generates an HTML table ()
      #
      # @param [Hash] options A hash with options
      # @return [String] A String containing the HTML code
      def to_html options={}
        options = {:style_with_inline_css=>false}.merge(options)
        builder = Nokogiri::XML::Builder.new do |doc|
          doc.table {
            self.each{|row|
              doc.tr {
                row.each {|cell|
                  classnames = cell.format.all_names.join(" ").strip
                  td_options = classnames != "" ? {:class=>classnames} : {}
                  td_options = td_options.merge({:style=>cell.format.to_css}) if options[:style_with_inline_css] and cell.format.to_css != ""
                  td_options = td_options.merge({:colspan=>cell.colspan}) if cell.colspan
                  td_options = td_options.merge({:rowspan=>cell.rowspan}) if cell.rowspan
                  unless cell.value.class == Workbook::NilValue
                    doc.td(td_options) {
                      doc.text cell.value
                    }
                  end
                }
              }
            }
          }
        end
        return builder.doc.to_xhtml
      end
    end
  end
end
