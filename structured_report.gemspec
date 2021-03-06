Gem::Specification.new do |s|
  s.name        = 'structured_report'
  s.version     = '1.0.7'
  s.summary     = "Simple structure for generating columned reports"
  s.description = ""
  s.authors     = ["Jason Smith"]
  s.email       = 'jsmith@red5studios.com'
  s.files       = Dir['lib/**/*.rb']
  
  s.add_runtime_dependency "writeexcel"
  
  s.add_development_dependency "rspec"
end