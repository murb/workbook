Gem::Specification.new do |s|
  s.name        = 'workbook'
  s.version     = '0.0.1'
  s.date        = '2012-01-02'
  s.summary     = "Workbook"
  s.description = "Workbook contains workbooks, as in a table, contains rows, contains cells"
  s.authors     = ["Maarten Brouwers"]
  s.email       = 'gem@murb.nl'
  s.files       = ["lib/workbook.rb"]
  s.add_dependency('spreadsheet', '>= 0.6.5')
  s.homepage    =
    'http://murb.nl/blog?tags=workbook'
end