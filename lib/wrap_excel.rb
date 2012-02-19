require "win32ole"
require File.join(File.dirname(__FILE__), 'wrap_excel/book')
require File.join(File.dirname(__FILE__), 'wrap_excel/sheet')
require File.join(File.dirname(__FILE__), 'wrap_excel/cell')
require File.join(File.dirname(__FILE__), 'wrap_excel/range')
require File.join(File.dirname(__FILE__), 'wrap_excel/cygwin') if RUBY_PLATFORM =~ /cygwin/
require "wrap_excel/version"

module WrapExcel

end
