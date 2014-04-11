module Workbook
  module Modules
    module Cache
      def fetch_cache(key, expires)
        @cache ||= {}
        if @cache[key]
          return @cache[key]
        else
          @cache[key] = yield
        end
        return @cache[key]
      end
    end
  end
end