module Workbook

  # Used in cases col or rowspans are used
  class NilValue

    # initialize this special nilvalue with a reason
    # @param [String] reason (currently only :covered, in case this cell is coverd because an adjecant cell spans over it)
    def initialize reason
      self.reason= reason
    end

    # returns the value of itself (nil)
    # @return [NilClass] nil
    def value
      nil
    end

    def <=> v
      value <=> v
    end

    def reason
      @reason
    end

    # set the reason why this value is nil
    def reason= reason
      if reason == :covered
        @reason = reason
      else
        raise "invalid reason given"
      end
    end

  end
end
