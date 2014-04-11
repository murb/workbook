# -*- encoding : utf-8 -*-
module Workbook
  module Modules
    # Adds simple caching
    module Cache
      # fetch cache, invalidates typically based on cache_valid_from call
      attr_accessor :debug_cache

      def cache_valid_from
        @cache_valid_from ||= Time.now
        @cache_valid_from
      end
      def invalidate_cache!
        @cache_valid_from = Time.now
        @cache_valid_from
      end
      def valid_cache_key?(key, expires=nil)
        cache_valid_from
        rv = (@cache[key] and (@cache[key][:inserted_at] > cache_valid_from) and (expires.nil? or @cache[key][:inserted_at] < expires)) ? true : false
        rv
      end
      def fetch_cache(key, expires=nil)
        @cache ||= {}
        if valid_cache_key?(key, expires)
          return @cache[key][:value]
        else
          @cache[key] = {
            value: yield,
            inserted_at: Time.now
          }
        end
        return @cache[key][:value]
      end
    end
  end
end