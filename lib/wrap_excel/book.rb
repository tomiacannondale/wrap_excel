# -*- coding: utf-8 -*-
module WrapExcel

  class Book
    attr_reader :book

    def initialize(file, options={ }, &block)
      options = {
        :read_only => true,
        :displayalerts => false,
        :visible => false
      }.merge(options)
      file = WrapExcel::Cygwin.cygpath('-w', file) if RUBY_PLATFORM =~ /cygwin/
      file = WIN32OLE.new('Scripting.FileSystemObject').GetAbsolutePathName(file)
      @winapp = WIN32OLE.new('Excel.Application')
      @winapp.DisplayAlerts = options[:displayalerts]
      @winapp.Visible = options[:visible]
      WIN32OLE.const_load(@winapp, WrapExcel) unless WrapExcel.const_defined?(:CONSTANTS)
      @book = @winapp.Workbooks.Open(file,{ 'ReadOnly' => options[:read_only] })

      if block
        begin
          yield self
        ensure
          close
        end
      end

      @book
    end

    def close
      @winapp.Workbooks.Close
      @winapp.Quit
    end

    def save
      @book.save
    end

    def [] sheet
      sheet += 1 if sheet.is_a? Numeric
      WrapExcel::Sheet.new(@book.Worksheets.Item(sheet))
    end

    def each
      @book.Worksheets.each do |sheet|
        yield WrapExcel::Sheet.new(sheet)
      end
    end

    def self.open(file, options={ }, &block)
      new(file, options, &block)
    end
  end
end
