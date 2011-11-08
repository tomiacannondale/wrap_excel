require 'bundler/gem_tasks'

require "rake"
require "rspec/core/rake_task"

RSpec::Core::RakeTask.new(:spec) do |s|
  s.rspec_opts = '-f d'
end
task :default => :spec
