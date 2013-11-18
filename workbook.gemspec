# -*- encoding: utf-8 -*-
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)

require "workbook/version"

Gem::Specification.new do |s|
  s.name        = 'workbook'
  s.rubyforge_project = 'workbook'
  s.version     = Workbook::VERSION
  s.date        = '2013-04-05'
  s.summary     = "Workbook is a datastructure to contain books of tables (an anlogy used in e.g. Excel)"
  s.description = "Workbook contains workbooks, as in a table, contains rows, contains cells, reads/writes excel, ods and csv and tab separated files, and offers basic diffing and sorting capabilities."
  s.authors     = ["Maarten Brouwers"]
  s.add_dependency('rubyzip')
  s.add_dependency('spreadsheet', '>= 0.7.5')
  s.add_dependency('fastercsv') if RUBY_VERSION < "1.9"
  s.add_dependency("rchardet", "~> 1.3")
  s.add_dependency("rake")
  s.add_dependency("json")
  s.add_dependency('rubyXL')
  if RUBY_VERSION < "1.9"
    s.add_dependency('nokogiri', "~> 1.5.10")
  else
    s.add_dependency('nokogiri')
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


