module Docxi
  module Util

    # Same rules as ActiveSupport's Object#blank?, which older versions of this
    # gem called without depending on ActiveSupport.
    def self.blank?(value)
      case value
      when nil, false then true
      when String then /\A[[:space:]]*\z/ === value
      else value.respond_to?(:empty?) ? !!value.empty? : false
      end
    end

    def self.present?(value)
      !blank?(value)
    end

  end
end
