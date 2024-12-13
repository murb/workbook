lib = File.expand_path("../lib", __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)

require "workbook/version"
require "date"

Gem::Specification.new do |s|
  s.name = "workbook"
  s.license = "MIT"
  s.version = Workbook::VERSION
  s.summary = "Workbook is a datastructure to contain books of tables (an anlogy used in e.g. Excel)"
  s.description = "Workbook contains workbooks, as in a table, contains rows, contains cells, reads/writes excel, ods and csv and tab separated files, and offers basic diffing and sorting capabilities."
  s.authors = ["Maarten Brouwers"]
  s.add_development_dependency("rake", "> 12.0")
  s.add_development_dependency("minitest", "> 5.4")
  s.add_development_dependency("byebug", "> 10")
  s.add_development_dependency("standard", "> 1.0")
  s.add_development_dependency("simplecov", "> 0.17.0")
  s.add_dependency("spreadsheet", "> 1.2")
  s.add_dependency("rchardet", ">= 1.8.0")
  s.add_dependency("json", "> 2.3")
  s.add_dependency("rubyzip", "> 1.2", ">= 1.2.1")
  s.add_dependency("caxlsx", "> 3.0")
  s.add_dependency("nokogiri", "> 1.10")
  s.add_dependency("csv", "> 3.0.0")

  s.platform = Gem::Platform::RUBY
  s.files = `git ls-files`.split($/)
  s.executables = s.files.grep(%r{^bin/}).map { |f| File.basename(f) }
  s.require_paths = ["lib"]
  s.email = ["gem@murb.nl"]
  s.homepage =
    "http://murb.nl/blog?tags=workbook"
end
