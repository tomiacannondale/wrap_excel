# -*- coding: utf-8 -*-
module WrapExcel
  class Sheet
    attr_reader :sheet
    include Enumerable

    def initialize(win32_worksheet)
      @sheet = win32_worksheet
    end

    def name
      @sheet.Name
    end

    def name= (new_name)
      @sheet.Name = new_name
    end

    def [] y, x
      yx = "#{y+1}_#{x+1}"
      @cells ||= { }
      @cells[yx] ||= WrapExcel::Cell.new(@sheet.Cells.Item(y+1, x+1))
    end

    def []= (y, x, value)
      @sheet.Cells.Item(y+1, x+1).Value = value
    end

    def each
      @sheet.UsedRange.Rows.each do |row_range|
        row_range.Cells.each do |cell|
          yield WrapExcel::Cell.new(cell)
        end
      end
    end

    def each_row(offset = 0)
      offset += 1
      @sheet.UsedRange.Rows.each do |row_range|
        next if row_range.row < offset
        yield WrapExcel::Range.new(row_range)
      end
    end

    def each_row_with_index(offset = 0)
      each_row(offset) do |row_range|
        yield WrapExcel::Range.new(row_range), (row_range.row - 1 - offset)
      end
    end

    def each_column(offset = 0)
      offset += 1
      @sheet.UsedRange.Columns.each do |column_range|
        next if column_range.column < offset
        yield WrapExcel::Range.new(column_range)
      end
    end

    def each_column_with_index(offset = 0)
      each_column(offset) do |column_range|
        yield WrapExcel::Range.new(column_range), (column_range.column - 1 - offset)
      end
    end

    def method_missing(id, *args)
      @sheet.send(id, *args)
    end
  end
end
