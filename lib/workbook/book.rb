# -*- encoding : utf-8 -*-
# frozen_string_literal: true
require 'open-uri'
require 'workbook/writers/xls_writer'
require 'workbook/writers/xlsx_writer'
require 'workbook/writers/html_writer'
require 'workbook/readers/xls_reader'
require 'workbook/readers/xls_shared'
require 'workbook/readers/xlsx_reader'
require 'workbook/readers/ods_reader'
require 'workbook/readers/csv_reader'
require 'workbook/readers/txt_reader'
require 'workbook/readers/txt_reader'
require 'workbook/modules/diff_sort'

module Workbook
  # The Book class is the container of sheets. It can be inialized by either the standard initalizer or the open method. The
  # Book class can also keep a reference to a template class, storing shared formatting options.
  #
  SUPPORTED_MIME_TYPES = [
    "application/zip",
    "text/plain",
    "application/x-ariadne-download",
    "application/vnd.ms-excel",
    "application/excel",
    "application/vnd.ms-office",
    "text/csv",
    "text/tab-separated-values",
    "application/x-ms-excel",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.oasis.opendocument.spreadsheet",
    "application/x-vnd.oasis.opendocument.spreadsheet",
    "CDF V2 Document, No summary info"
  ]

  class Book < Array


    include Workbook::Readers::XlsShared
    include Workbook::Writers::XlsWriter
    include Workbook::Writers::XlsxWriter
    include Workbook::Writers::HtmlWriter
    include Workbook::Readers::XlsReader
    include Workbook::Readers::OdsReader
    include Workbook::Readers::XlsxReader
    include Workbook::Readers::CsvReader
    include Workbook::Readers::TxtReader
    include Workbook::Modules::BookDiffSort

    # @param [Workbook::Sheet, Array] sheet    create a new workbook based on an existing sheet, or initialize a sheet based on the array
    # @return [Workbook::Book]
    def initialize sheet=nil
      if sheet.is_a? Workbook::Sheet
        self.push sheet
      elsif sheet
        self.push Workbook::Sheet.new(sheet, self, {})
      end
    end

    # @return [Workbook::Template] returns the template describing how the document should be/is formatted
    def template
      @template ||= Workbook::Template.new
    end

    # @param [Workbook::Format] template    a template describing how the document should be/is formatted
    def template= template
      raise ArgumentError, "format should be a Workboot::Format" unless template.is_a? Workbook::Template
      @template = template
    end

    # The title of the workbook
    #
    # @return [String] the title of the workbook
    def title
      (defined?(@title) and !@title.nil?) ? @title : "untitled document"
    end

    def title= t
      @title = t
    end

    # Push (like in array) a sheet to the workbook (parameter is optional, default is a new sheet)
    #
    # @param [Workbook::Sheet] sheet
    def push sheet=Workbook::Sheet.new
      super(sheet)
      sheet.book=(self)
    end

    # << (like in array) a sheet to the workbook (parameter is optional, default is a new sheet)
    #
    # @param [Workbook::Sheet] sheet
    def << sheet=Workbook::Sheet.new
      sheet = Workbook::Sheet.new(sheet) unless sheet.is_a? Workbook::Sheet
      super(sheet)
      sheet.book=(self)
    end


    # Sheet returns the first sheet of a workbook, or an empty one.
    #
    # @return [Workbook::Sheet] The first sheet, and creates an empty one if one doesn't exists
    def sheet
      push Workbook::Sheet.new unless first
      first
    end

    # If the first sheet has any contents
    #
    # @return [Boolean] returns true if the first sheet has contents
    def has_contents?
      sheet.has_contents?
    end

    # Loads an external file into an existing worbook
    #
    # @param [String] filename   a string with a reference to the file to be opened
    # @param [String] extension  an optional string enforcing a certain parser (based on the file extension, e.g. 'txt', 'csv' or 'xls')
    # @return [Workbook::Book] A new instance, based on the filename
    def import filename, extension=nil, options={}
      extension = file_extension(filename) unless extension
      if ['txt','csv','xml'].include?(extension)
        open_text filename, extension, options
      else
        open_binary filename, extension, options
      end
    end

    # Open the file in binary, read-only mode, do not read it, but pas it throug to the extension determined loaded
    #
    # @param [String] filename a string with a reference to the file to be opened
    # @param [String] extension an optional string enforcing a certain parser (based on the file extension, e.g. 'txt', 'csv' or 'xls')
    # @return [Workbook::Book] A new instance, based on the filename
    def open_binary filename, extension=nil, options={}
      extension = file_extension(filename) unless extension
      f = open(filename)
      send("load_#{extension}".to_sym, f, options)
    end

    # Open the file in non-binary, read-only mode, read it and parse it to UTF-8
    #
    # @param [String] filename   a string with a reference to the file to be opened
    # @param [String] extension  an optional string enforcing a certain parser (based on the file extension, e.g. 'txt', 'csv' or 'xls')
    def open_text filename, extension=nil, options={}
      extension = file_extension(filename) unless extension
      t = text_to_utf8(open(filename).read)
      send("load_#{extension}".to_sym, t, options)
    end

    # Writes the book to a file. Filetype is based on the extension, but can be overridden
    #
    # @param [String] filename   a string with a reference to the file to be written to
    # @param [Hash] options  depends on the writer chosen by the file's filetype
    def write filename, options={}
      extension = file_extension(filename)
      send("write_to_#{extension}".to_sym, filename, options)
    end


    # Helper method to convert text in a file to UTF-8
    #
    # @param [String] text a string to convert
    def text_to_utf8 text
      unless text.valid_encoding? and text.encoding == "UTF-8"
        # TODO: had some ruby 1.9 problems with rchardet ... but ideally it or a similar functionality will be reintroduced
        source_encoding = text.valid_encoding? ? text.encoding : "US-ASCII"
        text = text.encode('UTF-8', source_encoding, {:invalid=>:replace, :undef=>:replace, :replace=>""})
        text = text.gsub("\u0000","") # TODO: this cleanup of nil values isn't supposed to be needed...
      end
      text
    end

    # @param [String, File] filename   The full filename, or path
    #
    # @return [String] The file extension
    def file_extension(filename)
      ext = File.extname(filename).gsub('.','').downcase if filename
      # for remote files which has asset id after extension
      ext.split('?')[0]
    end

    # Load the CSV data contained in the given StringIO or String object
    #
    # @param [StringIO] stringio_or_string StringIO stream or String object, with data in CSV format
    # @param [Symbol] filetype (currently only :csv or :txt), indicating the format of the first parameter
    def read(stringio_or_string, filetype, options={})
      raise ArgumentError.new("The filetype parameter should be either :csv or :txt") unless [:csv, :txt].include?(filetype)
      t = stringio_or_string.respond_to?(:read) ? stringio_or_string.read : stringio_or_string.to_s
      t = text_to_utf8(t)
      send(:"parse_#{filetype}", t, options)
    end

    # Create or open the existing sheet at an index value
    #
    # @param [Integer] index    the index of the sheet
    def create_or_open_sheet_at index
      s = self[index]
      s = self[index] = Workbook::Sheet.new if s == nil
      s.book = self
      s
    end

    class << self
      # Create an instance from a file, using open.
      #
      # @param [String] filename of the document
      # @param [String] extension of the document (not required). The parser used is based on the extension of the file, this option allows you to override the default.
      # @return [Workbook::Book] A new instance, based on the filename
      def open filename, extension=nil
        wb = self.new
        wb.import filename, extension
        return wb
      end

      # Create an instance from the given stream or string, which should be in CSV or TXT format
      #
      # @param [StringIO] stringio_or_string StringIO stream or String object, with data in CSV or TXT format
      # @param [Symbol] filetype (currently only :csv or :txt), indicating the format of the first parameter
      # @return [Workbook::Book] A new instance
      def read stringio_or_string, filetype, options={}
        wb = self.new
        wb.read(stringio_or_string, filetype, options)
        wb
      end

    end
  end
end
