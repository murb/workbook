# -*- encoding : utf-8 -*-
module Workbook
	module Modules
    # Adds support for storing raw objects, used in e.g. Format and Template
		module RawObjectsStorage
      
	    # A raw is a 'raw' object, representing a workbook, or cell, or whatever... in a particular format (defined by its class)
      def add_raw raw_object
	      raws[raw_object.class]=raw_object
	    end
      
      # Returns true if there is a template for a certain class, otherwise false
	    def has_raw_for? raw_object_class
	      raws.each { |tc,t| return true if tc == raw_object_class}
        return false
	    end
      
      # Returns raw data stored for a type of raw object (if available)
      # @param [Class] raw_object_class (e.g. Spreadsheet::Format for the Spreadsheet-gem)
      def return_raw_for raw_object_class
	      raws.each { |tc,t| return t if tc == raw_object_class}
        return nil
      end 
      
      # Remove all raw data references
      def remove_all_raws!
        @raws = {}
      end
      
      # Return all raw data references
	    def raws
        @raws = {} unless defined? @raws
        @raws
	    end
		end
	end
end
