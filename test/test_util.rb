require 'test_helper'

class TestUtil < Minitest::Test
  BLANK = [nil, false, '', ' ', "\t\r\n", " 　", [], {}, :''].freeze
  PRESENT = [true, 0, 1.5, 'a', ' a ', [nil], { a: 1 }, :a, Object.new, Time.at(0)].freeze

  def test_blank_values
    BLANK.each do |value|
      assert Docxi::Util.blank?(value), "#{value.inspect} should be blank"
      refute Docxi::Util.present?(value), "#{value.inspect} should not be present"
    end
  end

  def test_present_values
    PRESENT.each do |value|
      refute Docxi::Util.blank?(value), "#{value.inspect} should not be blank"
      assert Docxi::Util.present?(value), "#{value.inspect} should be present"
    end
  end

  def test_matches_activesupport_when_it_is_loaded
    skip 'ActiveSupport is not loaded in this run' unless ''.respond_to?(:blank?)

    (BLANK + PRESENT).each do |value|
      assert_equal value.blank?, Docxi::Util.blank?(value), "blank? disagrees for #{value.inspect}"
    end
  end
end
