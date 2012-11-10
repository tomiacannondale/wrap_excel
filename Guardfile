# -*- mode: ruby -*-
# A sample Guardfile
# More info at https://github.com/guard/guard#readme

guard 'rspec', cli: "--color", all_after_pass: false, all_on_start: false do
  watch(%r{^spec/.+_spec\.rb$})
  watch(%r{^lib/wrap_excel/(.+)\.rb$})     { |m| "spec/#{m[1]}_spec.rb" }
  watch('lib/wrap_excel.rb')     { "spec" }
  watch('spec/spec_helper.rb')  { "spec" }
end
