module Workbook
	module Modules
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
      
      def return_raw_for raw_object_class
	      raws.each { |tc,t| return t if tc == raw_object_class}
        return nil
      end 
      
      def remove_all_raws!
        @raws = {}
      end
      
	    def raws
        @raws = {} unless defined? @raws
        @raws
	    end
		end
	end
end