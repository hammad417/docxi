require_relative 'lib/docxi/version'

Gem::Specification.new do |s|
  s.name        = 'docxi'
  s.version     = Docxi::VERSION
  s.summary     = 'Docx Template Generator'
  s.description = 'Create docx template using Ruby'
  s.authors     = ['Hammad Maqbool']
  s.email       = ['hammad.maqbool@phaedrasolutions.com', 'hammad417@gmail.com']
  s.homepage    = 'https://github.com/hammad417/docxi'
  s.license     = 'MIT'

  s.metadata = {
    'source_code_uri'       => 'https://github.com/hammad417/docxi',
    'changelog_uri'         => 'https://github.com/hammad417/docxi/blob/master/CHANGELOG.md',
    'bug_tracker_uri'       => 'https://github.com/hammad417/docxi/issues',
    'rubygems_mfa_required' => 'true'
  }

  s.required_ruby_version = '>= 2.5'

  s.files = Dir['lib/**/*'] + ['LICENSE', 'Rakefile', 'README.md', 'CHANGELOG.md']

  s.add_dependency 'nokogiri', '>= 1.8', '< 2'
  s.add_dependency 'rubyzip', '>= 1.3', '< 4'

  s.add_development_dependency 'minitest', '>= 5.15', '< 7'
  s.add_development_dependency 'rake', '>= 12', '< 14'
end
