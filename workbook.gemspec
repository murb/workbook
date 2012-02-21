# -*- encoding: utf-8 -*-
$:.push File.expand_path("../lib", __FILE__)
require "workbook/version"


spec = Gem::Specification.new do |s|
  s.name        = 'workbook'
  s.version     = Workbook::VERSION
  s.date        = '2012-01-02'
  s.summary     = "Workbook"
  s.description = "Workbook contains workbooks, as in a table, contains rows, contains cells"
  s.authors     = ["Maarten Brouwers"]
  s.platform    = Gem::Platform::RUBY
  s.files         = `git ls-files`.split("\n")
  s.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")
  s.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ["lib"]
  s.email       = 'gem@murb.nl'
  s.add_dependency('spreadsheet', '>= 0.6.7')
  s.homepage    =
    'http://murb.nl/blog?tags=workbook'
end

if $0 == __FILE__
   Gem.manage_gems
   Gem::Builder.new(spec).build
end
