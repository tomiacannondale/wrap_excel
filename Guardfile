# -*- mode: ruby -*-
# A sample Guardfile
# More info at https://github.com/guard/guard#readme

guard 'rspec', :version => 2 do
  watch(%r{^spec/.+_spec\.rb$})
  watch(%r{^lib/wrap_excel/(.+)\.rb$})     { |m| "spec/#{m[1]}_spec.rb" }
  watch('lib/wrap_excel.rb')     { "spec" }
  watch('spec/spec_helper.rb')  { "spec" }
end
