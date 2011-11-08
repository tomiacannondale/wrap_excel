# -*- coding: utf-8 -*-
module WrapExcel
  class Range
    def initialize(win32_range)
      @range = win32_range
    end

    def each
      @range.each do |row_or_column|
        yield WrapExcel::Cell.new(row_or_column)
      end
    end

    def values(range = nil)
      if range
        min = range.min + 1
        max = range.max + 1
        result = @range.Range(@range.Cells.Item(min), @range.Cells(max)).value
        result.is_a?(Array) ? result[0] : [result]
      else
        @range.Cells.value.flatten
      end
    end

    def [] index
      @cells ||= []
      @cells[index + 1] ||= WrapExcel::Cell.new(@range.Cells.Item(index + 1))
    end

    def method_missing(id, *args)
      @range.send(id, *args)
    end
  end
end
