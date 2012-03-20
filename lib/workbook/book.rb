require 'workbook/writers/xls_writer'
require 'workbook/readers/xls_reader'
require 'workbook/readers/csv_reader'
require 'workbook/readers/txt_reader'
require 'rchardet'

module Workbook
  class Book < Array
    include Workbook::Writers::XlsWriter
    include Workbook::Readers::XlsReader
    include Workbook::Readers::CsvReader
    include Workbook::Readers::TxtReader    
    
    attr_accessor :title
    attr_accessor :template
    attr_accessor :default_rewrite_header
    
    # def initialize sheet=Workbook::Sheet.new
    #   push sheet if sheet
    # end
    
    def initialize sheet=Workbook::Sheet.new([], self, options={})
      if sheet.is_a? Workbook::Sheet
        push sheet
      else
        push Workbook::Sheet.new(sheet, self, options)
      end
    end
    
    def template
      @template ||= Workbook::Template.new
    end
    
    def template= template
      raise ArgumentError, "format should be a Workboot::Format" unless template.is_a? Workbook::Template
      @template = template
    end
    
    def title
      @title ? @title : "untitled document"
    end
    
  	def push sheet=Workbook::Sheet.new
      super(sheet)
    end
    
    # Returns the first sheet, and creates an empty one if one doesn't exists.
    def sheet
      push Workbook::Sheet.new unless first
      first
    end
    
    def has_contents?
      sheet.has_contents?
    end
    
    # Loads an external file into an existing worbook
    def open filename, ext=nil
      ext = file_extension(filename) unless ext
      if ['txt','csv','xml'].include?(ext)
        open_text filename, ext
      else
        open_binary filename, ext
      end
    end
    
    # open the file in binary, read-only mode, do not read it, but pas it throug to the extension determined loaded
    def open_binary filename, ext=nil
      ext = file_extension(filename) unless ext
      f = File.open(filename,'rb')
      send("load_#{ext}".to_sym,f)
    end
    
    # open the file in non-binary, read-only mode, read it and parse it to UTF-8
    def open_text filename, ext=nil
      ext = file_extension(filename) unless ext
      f = File.open(filename,'r')
      t = f.read
      detected_encoding = CharDet.detect(t)['encoding']
      t = Iconv.conv("UTF-8//TRANSLIT//IGNORE",detected_encoding,t)
      send("load_#{ext}".to_sym,t)
    end
    
    def file_extension(filename)
      File.extname(filename).gsub('.','')
    end
    
    # Create an instance from a file, using open.
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
    
    def default_rewrite_header?
      return true if default_rewrite_header.nil?
      default_rewrite_header
    end
    
  end
end
