# -*- encoding : utf-8 -*-
$KCODE="u" if RUBY_VERSION < "1.9"
require_relative 'workbook/modules/cache'
require_relative 'workbook/book'
require_relative 'workbook/sheet'
require_relative 'workbook/table'
require_relative 'workbook/nil_value'
require_relative 'workbook/row'
require_relative 'workbook/cell'
require_relative 'workbook/format'
require_relative 'workbook/template'
require_relative 'workbook/version'
require_relative 'workbook/column'

module Workbook
  class << self
    def caching_enabled?
      @@caching_enabled ||= true
    end
    # Disable caching globally (wherever applicable)
    def disable_caching!
      @@caching_enabled = false
    end
    # Enable caching globally (wherever applicable)
    def enable_caching!
      @@caching_enabled = true
    end
  end
end
