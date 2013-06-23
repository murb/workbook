# -*- encoding : utf-8 -*-
require 'json'

module Workbook
  module Writers
    module JsonTableWriter
      # Output the current workbook to JSON format
      #
      # @param [String] filename
      # @param [Hash] options  
      # @return [String] csv (comma separated values in a string)
      def to_json options={}
        JSON.generate(to_array_of_hashes_with_values(options))
      end
      
      # Output the current workbook to JSON format
      #
      # @param [String] filename
      # @param [Hash] options  
      # @return [String] csv (comma separated values in a string)
      def to_array_of_hashes_with_values options={}
        self.collect{|a| a.to_hash_with_values}
      end
      
      # Write the current workbook to JSON format
      #
      # @param [String] filename
      # @param [Hash] options   see #to_json
      # @return [String] filename
      def write_to_json filename="#{title}.json", options={}
        File.open(filename, 'w') {|f| f.write(to_json(options)) }
        return filename
      end

    end
  end
end
