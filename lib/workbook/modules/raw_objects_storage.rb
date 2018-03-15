# -*- encoding : utf-8 -*-
# frozen_string_literal: true
module Workbook
  module Modules
    # Adds support for storing raw objects, used in e.g. Format and Template
    module RawObjectsStorage

      # A raw is a 'raw' object, representing a workbook, or cell, or whatever... in a particular format (defined by its class)
      def add_raw raw_object, options={}
        class_of_obj = options[:raw_object_class] ? options[:raw_object_class] : raw_object.class
        raws[class_of_obj]=raw_object
      end

      # Returns true if there is a template for a certain class, otherwise false
      def has_raw_for? raw_object_class
        available_raws.include? raw_object_class
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

      # Lists the classes for which raws are available
      # @return [Array<Object>] array with the classes available
      def available_raws
        raws.keys
      end

      # Return all raw data references
      def raws
        @raws = {} unless defined? @raws
        @raws
      end
    end
  end
end
