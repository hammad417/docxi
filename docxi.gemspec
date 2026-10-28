Gem::Specification.new do |s|
  s.name        = 'docxi'
  s.version     = '0.0.6'
  s.date        = '2016-10-14'
  s.summary     = 'Docx Template Generator'
  s.description = 'Create docx template using Ruby'
  s.authors     = ["Irfan Babar"]
  s.email       = 'salman.shan@phaedrasolutions.com'
  s.homepage    = ''
  s.license     = 'MIT'

  s.required_ruby_version = '>= 2.5'

  s.files = Dir["lib/**/*"] + ["LICENSE", "Rakefile", "README.md"]

  s.add_dependency 'nokogiri', '>= 1.8', '< 2'
  s.add_dependency 'rubyzip', '>= 1.3', '< 4'

  s.add_development_dependency 'minitest', '>= 5.15', '< 7'
  s.add_development_dependency 'rake', '>= 12', '< 14'
end
