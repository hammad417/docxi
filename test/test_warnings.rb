require 'test_helper'

class TestWarnings < Minitest::Test
  LIB = File.expand_path('../lib', __dir__)

  # On Ruby 3.4+ this also catches mutated string literals, which will break
  # once Ruby freezes them by default.
  def test_rendering_emits_no_deprecation_warnings_from_docxi
    skip 'warning categories need Ruby 2.7+' unless Warning.respond_to?(:[]=)

    previous = Warning[:deprecated]
    Warning[:deprecated] = true
    _out, err = capture_io { SampleDocuments.all.each_value(&:render) }

    assert_empty err.lines.select { |line| line.include?(LIB) }
  ensure
    Warning[:deprecated] = previous if Warning.respond_to?(:[]=)
  end
end
