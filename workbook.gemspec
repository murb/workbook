# -*- encoding: utf-8 -*-
$:.push File.expand_path("../lib", __FILE__)
require "workbook"


Gem::Specification.new do |s|
  s.name        = 'workbook'
  s.rubyforge_project = 'workbook'
  s.version     = '0.0.25'
  s.date        = '2012-02-23'
  s.summary     = "Workbook is a datastructure to contain books of tables (an anlogy used in e.g. Excel)"
  s.description = "Workbook contains workbooks, as in a table, contains rows, contains cells, reads/writes excels and csv's and tab separated, and offers basic diffing and sorting capabilities."
  s.authors     = ["Maarten Brouwers"]
  s.add_dependency('spreadsheet', '>= 0.6.8')
  s.add_dependency('fastercsv')
  s.add_dependency("rchardet", "~> 1.3")
  s.platform    = Gem::Platform::RUBY
  s.files         = `git ls-files`.split("\n")
  s.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")
  s.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ["lib"]
  s.email       = ['gem@murb.nl']
  s.homepage    =
    'http://murb.nl/blog?tags=workbook'
end


