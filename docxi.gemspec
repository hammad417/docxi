Gem::Specification.new do |s|
  s.name        = 'docxi'
  s.version     = '0.0.3'
  s.date        = '2016-10-14'
  s.summary     = 'Docx Template Generator'
  s.description = 'Create docx template using Ruby'
  s.authors     = ["Irfan Babar"]
  s.email       = 'salman.shan@phaedrasolutions.com'
  s.homepage    = ''
  s.license     = 'MIT'

  s.files = Dir["lib/**/*"] + ["LICENSE", "Rakefile", "README.md"]

  s.add_dependency 'nokogiri', '>= 1.6.3', '< 1.7'
  s.add_dependency 'rubyzip', '~> 0.9.9'
end
