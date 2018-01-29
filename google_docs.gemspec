require File.expand_path('../lib/google_docs/version', __FILE__)

Gem::Specification.new do |spec|
  spec.name          = 'google_docs'
  spec.version       = GoogleDocs::VERSION
  spec.authors       = ['JetRuby']
  spec.email         = ['engineering@jetruby.com']

  spec.summary       = 'Google Docs'
  spec.description   = 'A simple gem to create styles for google document'
  spec.homepage      = 'http://jetruby.com'
  spec.license       = 'MIT'
  spec.files         = `git ls-files -z`.split("\x0").reject do |f|
    f.match(%r{^(test|spec|features)/})
  end
  spec.bindir        = 'exe'
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ['lib']

  spec.add_development_dependency 'bundler', '~> 1.14'
  spec.add_development_dependency 'rake', '~> 10.0'
  spec.add_development_dependency 'rspec', '~> 3.5'
end
