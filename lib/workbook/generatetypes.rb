# frozen_string_literal: true
["Numeric","String","Time","Date","TrueClass","FalseClass","NilClass"].each do |type|
  f = File.open(File.join(File.dirname(__FILE__),"types","#{type}.rb"),'w+')
  puts f.inspect
  doc="require 'workbook/cell'

module Workbook
  module Types
    class #{type} < #{type}
      include Workbook::Cell
    end
  end
end"
f.write(doc)
end