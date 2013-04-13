# -*- encoding : utf-8 -*-
require 'workbook/writers/xls_writer'
require 'workbook/readers/xls_reader'
require 'workbook/readers/xls_shared'
require 'workbook/readers/xlsx_reader'
require 'workbook/readers/ods_reader'
require 'workbook/readers/csv_reader'
require 'workbook/readers/txt_reader'

module Workbook
  # The Book class is the container of sheets. It can be inialized by either the standard initalizer or the open method. The 
  # Book class can also keep a reference to a template class, storing shared formatting options.
  # 
  class Book < Array
    include Workbook::Readers::XlsShared
    include Workbook::Writers::XlsWriter
    include Workbook::Readers::XlsReader
    include Workbook::Readers::OdsReader
    include Workbook::Readers::XlsxReader
    include Workbook::Readers::CsvReader
    include Workbook::Readers::TxtReader    
    
    attr_accessor :title
    attr_accessor :template
    
    # @param [Workbook::Sheet, Array] sheet    create a new workbook based on an existing sheet, or initialize a sheet based on the array
    # @return [Workbook::Book]  
    def initialize sheet=Workbook::Sheet.new([], self, options={})
      if sheet.is_a? Workbook::Sheet
        push sheet
      else
        push Workbook::Sheet.new(sheet, self, options)
      end
    end
    
    # @return [Workbook::Format] returns the template describing how the document should be/is formatted
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
      @title ? @title : "untitled document"
    end
    
    # Push (like in array) a sheet to the workbook (parameter is optional, default is a new sheet)
    #
    # @param [Workbook::Sheet] sheet
  	def push sheet=Workbook::Sheet.new
      super(sheet)
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
    def open filename, extension=nil
      extension = file_extension(filename) unless extension
      if ['txt','csv','xml'].include?(extension)
        open_text filename, extension
      else
        open_binary filename, extension
      end
    end
    
    # Open the file in binary, read-only mode, do not read it, but pas it throug to the extension determined loaded
    #
    # @param [String] filename a string with a reference to the file to be opened
    # @param [String] extension an optional string enforcing a certain parser (based on the file extension, e.g. 'txt', 'csv' or 'xls')
    # @return [Workbook::Book] A new instance, based on the filename
    def open_binary filename, extension=nil
      extension = file_extension(filename) unless extension
      f = File.open(filename,'rb')
      send("load_#{extension}".to_sym,f)
    end
    
    # Open the file in non-binary, read-only mode, read it and parse it to UTF-8
    #
    # @param [String] filename   a string with a reference to the file to be opened
    # @param [String] extension  an optional string enforcing a certain parser (based on the file extension, e.g. 'txt', 'csv' or 'xls')
    def open_text filename, extension=nil
      extension = file_extension(filename) unless extension
      f = File.open(filename,'r')
      t = f.read
      t = text_to_utf8(t)
      send("load_#{extension}".to_sym,t)
    end
    
    
    # Helper method to convert text in a file to UTF-8
    # 
    # @param [String] text a string to convert
    def text_to_utf8 text
      if RUBY_VERSION < '1.9'
        require 'rchardet' 
        require 'iconv'
        detected_encoding = CharDet.detect(text)
        detected_encoding = detected_encoding['encoding']
        text = Iconv.conv("UTF-8//TRANSLIT//IGNORE",detected_encoding,text)
        text = text.gsub("\xEF\xBB\xBF", '') # removing the BOM...
      else
        unless text.valid_encoding? and text.encoding == "UTF-8"
          # TODO: had some ruby 1.9 problems with rchardet ... but ideally it or a similar functionality will be reintroduced
          source_encoding = text.valid_encoding? ? text.encoding : "US-ASCII"
          text = text.encode('UTF-8', source_encoding, {:invalid=>:replace, :undef=>:replace, :replace=>""})
          text = text.gsub("\u0000","") # TODO: this cleanup of nil values isn't supposed to be needed...
        end
      end
      return text
    end
    
    # @param [String] filename   The full filename, or path
    #
    # @return [String] The file extension
    def file_extension(filename)
      File.extname(filename).gsub('.','').downcase if filename
    end
    
    # Create an instance from a file, using open.
    #
    # @param [String] filename of the document
    # @param [String] extension of the document (not required). The parser used is based on the extension of the file, this option allows you to override the default.
    # @return [Workbook::Book] A new instance, based on the filename
    def self.open filename, extension=nil
      wb = self.new
      wb.open filename, extension
      return wb
    end

    # Create an instance from the given stream or string, which should be in CSV or TXT format
    #
    # @param [StringIO] a StringIO stream or String object, with data in CSV or TXT format
    # @param [Symbol] :csv or :txt, indicating the format of the first parameter
    # @return [Workbook::Book] A new instance
    def self.read(stringio_or_string, filetype)
      wb = self.new
      wb.read(stringio_or_string, filetype)
      wb
    end
    
    # Load the CSV data contained in the given StringIO or String object
    #
    # @param [StringIO] a StringIO stream or String object, with data in CSV format
    # @param [Symbol] :csv or :txt, indicating the format of the first parameter
    def read(stringio_or_string, filetype)
      raise ArgumentError.new("The filetype parameter should be either :csv or :txt") unless [:csv, :txt].include?(filetype)
      t = stringio_or_string.respond_to?(:read) ? stringio_or_string.read : stringio_or_string.to_s
      t = text_to_utf8(t)
      send(:"parse_#{filetype}", t)
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
    
  end
end
