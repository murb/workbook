require 'workbook/writers/xls_writer'
require 'workbook/readers/xls_reader'
require 'workbook/readers/xlsx_reader'
require 'workbook/readers/csv_reader'
require 'workbook/readers/txt_reader'
require 'rchardet'

module Workbook
  # The Book class is the container of sheets. It can be inialized by either the standard initalizer or the open method. The 
  # Book class can also keep a reference to a template class, storing shared formatting options.
  # 
  class Book < Array
    include Workbook::Writers::XlsWriter
    include Workbook::Readers::XlsReader
    include Workbook::Readers::XlsxReader
    include Workbook::Readers::CsvReader
    include Workbook::Readers::TxtReader    
    
    attr_accessor :title
    attr_accessor :template
    
    # @param [Workbook::Sheet, Array] Create a new workbook based on an existing sheet, or initialize a sheet based on the array
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
    
    # @param [Workbook::Format] a template describing how the document should be/is formatted
    def template= template
      raise ArgumentError, "format should be a Workboot::Format" unless template.is_a? Workbook::Template
      @template = template
    end
    
    # @return [String] the title of the workbook
    def title
      @title ? @title : "untitled document"
    end
    
  	def push sheet=Workbook::Sheet.new
      super(sheet)
    end
    
    # @return [Workbook::Sheet] The first sheet, and creates an empty one if one doesn't exists
    def sheet
      push Workbook::Sheet.new unless first
      first
    end
    
    # @return [Boolean] returns true if the first sheet has contents
    def has_contents?
      sheet.has_contents?
    end
    
    # Loads an external file into an existing worbook
    #
    # @param [String] a string with a reference to the file to be opened
    # @param [String] an optional string enforcing a certain parser (based on the file extension, e.g. 'txt', 'csv' or 'xls')
    # @return [Workbook::Book] A new instance, based on the filename
    def open filename, ext=nil
      ext = file_extension(filename) unless ext
      if ['txt','csv','xml'].include?(ext)
        open_text filename, ext
      else
        open_binary filename, ext
      end
    end
    
    # Open the file in binary, read-only mode, do not read it, but pas it throug to the extension determined loaded
    #
    # @param [String] a string with a reference to the file to be opened
    # @param [String] an optional string enforcing a certain parser (based on the file extension, e.g. 'txt', 'csv' or 'xls')
    # @return [Workbook::Book] A new instance, based on the filename
    def open_binary filename, ext=nil
      ext = file_extension(filename) unless ext
      f = File.open(filename,'rb')
      send("load_#{ext}".to_sym,f)
    end
    
    # Open the file in non-binary, read-only mode, read it and parse it to UTF-8
    #
    # @param [String] a string with a reference to the file to be opened
    # @param [String] an optional string enforcing a certain parser (based on the file extension, e.g. 'txt', 'csv' or 'xls')
    def open_text filename, ext=nil
      ext = file_extension(filename) unless ext
      f = File.open(filename,'r')
      t = f.read
      detected_encoding = CharDet.detect(t)['encoding']
      t = Iconv.conv("UTF-8//TRANSLIT//IGNORE",detected_encoding,t)
      send("load_#{ext}".to_sym,t)
    end
    
    # @param [String] The full filename, or path
    #
    # @return [String] The file extension
    def file_extension(filename)
      File.extname(filename).gsub('.','').downcase if filename
    end
    
    # Create an instance from a file, using open.
    #
    # @param [String] the filename of the document
    # @param [String] (not required) enforce a certain extension, the parser is based on the extension of the file
    # @return [Workbook::Book] A new instance, based on the filename
    def self.open filename, ext=nil
      wb = self.new
      wb.open filename, ext
      return wb
    end

    def create_or_open_sheet_at index
      s = self[index]
      s = self[index] = Workbook::Sheet.new if s == nil
      s.book = self
      s 
    end
    
    def sort
      raise Exception("Books can't be sorted")
    end

    
  end
end
