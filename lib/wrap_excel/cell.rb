# -*- coding: utf-8 -*-

module WrapExcel
  class Cell
    attr_reader :cell

    def initialize(win32_cell)
      if win32_cell.MergeCells
        @cell = win32_cell.MergeArea.Item(1,1)
      else
        @cell = win32_cell
      end
    end

    def method_missing(id, *args)
      @cell.send(id, *args)
    end
  end
end
