# -*- encoding : utf-8 -*-
module Workbook
  module Modules
    # Adds simple caching
    module Cache
      attr_accessor :debug_cache

      # Caching enabled?
      # @return [Boolean]
      def caching_enabled?
        Workbook.caching_enabled?
      end

      # Return valid cache time, if caching is enabled, otherwise calls #invalidate_cache!
      # @return [Time] Timestamp after which cache is valid
      def cache_valid_from
        if caching_enabled?
          @cache_valid_from ||= Time.now
        else
          invalidate_cache!
        end
        @cache_valid_from
      end

      # Invalidate all caches on this instance, and reset
      # @return [Time] Timestamp after which cache is valid (=current time, hence everything stored before is invalid)
      def invalidate_cache!
        @cache_valid_from = Time.now
      end

      # Check if currently stored key is available and still valid
      # @return [Boolean]
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