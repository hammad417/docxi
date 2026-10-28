# DOCXI_WITH_RAILS=1 loads the full Rails stack first, as a Rails app would,
# to prove docxi produces the same output next to it.
WITH_RAILS = ENV['DOCXI_WITH_RAILS'] == '1'

if WITH_RAILS
  require 'logger' # Rails <= 7.0 needs this with concurrent-ruby >= 1.3.5
  require 'rails/all'
  require 'active_support/all'
end

require 'minitest/autorun'
require 'stringio'
require 'docxi'
require 'support/package'
require 'support/sample_documents'

module DocxiAssertions
  W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'.freeze
  V = 'urn:schemas-microsoft-com:vml'.freeze

  def assert_same_xml(expected, actual, label)
    expected = Package.canonical(expected)
    actual = Package.canonical(actual)
    assert expected == actual, lambda {
      at = (0...[expected.size, actual.size].min).find { |i| expected[i] != actual[i] } || [expected.size, actual.size].min
      from = [at - 80, 0].max
      "#{label} differs at character #{at}\n  expected: #{expected[from, 160]}\n  actual:   #{actual[from, 160]}"
    }
  end

  def render_parts
    word = Docxi::Document.new
    yield word.document
    Package.parts(word.render.string)
  end

  def xml(data)
    Nokogiri::XML(data)
  end
end

Minitest::Test.include(DocxiAssertions)
