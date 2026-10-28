require 'test_helper'

# test/golden was rendered by the untouched 0.0.6 code (with ActiveSupport
# loaded, as in a Rails app). Every part must still match, compared as
# canonical XML so different nokogiri versions don't cause false failures.
# After an intentional output change, regenerate with `rake golden`.
class TestGoldenOutput < Minitest::Test
  GOLDEN = File.expand_path('golden', __dir__)

  %w[kitchen_sink blank_options].each do |name|
    define_method("test_#{name}_matches_golden_files") do
      parts = Package.parts(SampleDocuments.public_send(name).render.string)

      assert_equal File.read(File.join(GOLDEN, name, 'entries.txt')).lines.map(&:chomp), parts.keys

      parts.each do |part, data|
        expected = File.binread(File.join(GOLDEN, name, part))
        if part.end_with?('.xml', '.rels')
          assert_same_xml expected, data, "#{name}/#{part}"
        else
          assert expected == data.b, "#{name}/#{part} differs"
        end
      end
    end
  end
end
