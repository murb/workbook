# -*- encoding: utf-8 -*-
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)

require "workbook/version"
require 'date'

Gem::Specification.new do |s|
  s.name        = 'workbook'
  s.rubyforge_project = 'workbook'
  s.license = "MIT"
  s.version     = Workbook::VERSION
  s.date        = Time.new.to_date.to_s
  s.summary     = "Workbook is a datastructure to contain books of tables (an anlogy used in e.g. Excel)"
  s.description = "Workbook contains workbooks, as in a table, contains rows, contains cells, reads/writes excel, ods and csv and tab separated files, and offers basic diffing and sorting capabilities."
  s.authors     = ["Maarten Brouwers"]
  s.add_development_dependency 'ruby-prof', '~> 0.14'
  s.add_dependency('rubyzip', '~>1')
  s.add_dependency('spreadsheet', '~> 1.0')
  s.add_development_dependency('minitest', '~> 5.4')
  s.add_dependency('fastercsv') if RUBY_VERSION < "1.9"
  s.add_dependency("rchardet", "~> 1.3")
  s.add_dependency("rake", '~> 10.0')
  s.add_dependency("json", '~> 1.8')
  s.add_dependency("zip-zip", '~> 0.2') #actually a hack...
  #s.add_dependency('rubyXL', '~> 3.3')
  s.add_dependency('roo', '~> 1.13')
  s.add_dependency('axlsx', '~> 2.0.1')
  if RUBY_VERSION < "1.9"
    s.add_dependency('nokogiri', "~> 1.5.10")
  else
    s.add_dependency('nokogiri', '~> 1.6')
  end
  s.platform    = Gem::Platform::RUBY
  s.files         = `git ls-files`.split($/)
  s.test_files    = s.files.grep(%r{^(test|spec|features)/})
  s.executables   = s.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  s.require_paths = ["lib"]
  s.email       = ['gem@murb.nl']
  s.homepage    =
    'http://murb.nl/blog?tags=workbook'
end


