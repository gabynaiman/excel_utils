# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require_relative 'lib/excel_utils/version'

Gem::Specification.new do |spec|
  spec.name          = 'excel_utils'
  spec.version       = ExcelUtils::VERSION
  spec.authors       = ['Gabriel Naiman']
  spec.email         = ['gabynaiman@gmail.com']
  spec.summary       = 'Excel utils for easy read and write'
  spec.description   = 'Excel utils for easy read and write'
  spec.homepage      = 'https://github.com/gabynaiman/excel_utils'
  spec.license       = 'MIT'

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ['lib']

  spec.add_runtime_dependency 'inflecto', '~> 0.0'
  spec.add_runtime_dependency 'roo', '~> 2.8'
  spec.add_runtime_dependency 'roo-xls', '~> 1.2'
  spec.add_runtime_dependency 'write_xlsx', '~> 0.85'
  spec.add_runtime_dependency 'nesquikcsv', '~> 0.1'

  spec.add_development_dependency 'rake', '~> 12.0'
  spec.add_development_dependency 'minitest', '~> 5.0', '< 5.11'
  spec.add_development_dependency 'minitest-colorin', '~> 0.1'
  spec.add_development_dependency 'minitest-line', '~> 0.6'
  spec.add_development_dependency 'simplecov', '~> 0.12'
  spec.add_development_dependency 'coveralls', '~> 0.8'
  spec.add_development_dependency 'pry-nav', '~> 0.2'
end