require 'test_helper'

class TestRailsIndependence < Minitest::Test
  def test_works_without_activesupport
    skip 'this run loads Rails on purpose' if WITH_RAILS

    refute defined?(ActiveSupport), 'something loaded ActiveSupport'
    refute ''.respond_to?(:blank?), 'something added String#blank?'
    assert_includes Package.parts(SampleDocuments.blank_options.render.string).keys, 'word/document.xml'
  end

  def test_rails_is_really_loaded_when_requested
    skip 'this run does not load Rails' unless WITH_RAILS

    assert defined?(Rails::VERSION), 'Rails was not loaded'
    assert ''.respond_to?(:blank?), 'ActiveSupport core extensions were not loaded'
  end
end
