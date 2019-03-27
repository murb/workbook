# frozen_string_literal: true

require 'rake'
require 'rake/testtask'
require 'bundler'
require "bundler/gem_tasks"

#Bundler::GemHelper.install_tasks

task :default => [:test]

Rake::TestTask.new do |t|
  t.libs << "test"
  t.test_files = FileList['test/test*.rb','test/writers/test*.rb']
  t.verbose = false
end
