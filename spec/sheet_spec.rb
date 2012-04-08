# -*- coding: utf-8 -*-
require File.join(File.dirname(__FILE__), './spec_helper')

describe WrapExcel::Sheet do
  before do
    @dir = create_tmpdir
    @book = WrapExcel::Book.open(@dir + '/simple.xls')
    @sheet = @book[0]
  end

  after do
    @book.close
    rm_tmp(@dir)
  end

  describe ".initialize" do
    context "when open sheet protected(with password is 'protect')" do
      before do
        @book_protect = WrapExcel::Book.new(@dir + '/protected_sheet.xls', :visible => true)
        @protected_sheet = @book_protect['protect']
      end

      after do
        @book_protect.close
      end

      it { @protected_sheet.ProtectContents.should be_true }

      it "protected sheet can't be write" do
        expect { @protected_sheet[0,0] = 'write' }.to raise_error
      end
    end

  end

  shared_context "sheet 'open book with blank'" do
    before do
      @book_with_blank = WrapExcel::Book.open(@dir + '/book_with_blank.xls')
      @sheet_with_blank = @book_with_blank[0]
    end

    after do
      @book_with_blank.close
    end
  end

  describe "access sheet name" do
    describe "#name" do
      it 'get sheet1 name' do
        @sheet.name.should eq 'Sheet1'
      end
    end

    describe "#name=" do
      it 'change sheet1 name to foo' do
        @sheet.name = 'foo'
        @sheet.name.should eq 'foo'
      end
    end
  end

  describe 'access cell' do
    describe "#[]" do
      context "access [0,0]" do
        it { @sheet[0, 0].should be_kind_of WrapExcel::Cell }
        it { @sheet[0, 0].value.should eq 'simple' }
      end

      context "access [0, 0], [0, 1], [2, 0]" do
        it "should get every values" do
          @sheet[0, 0].value.should eq 'simple'
          @sheet[0, 1].value.should eq 'workbook'
          @sheet[2, 0].value.should eq 'matz'
        end
      end
    end

    it "change a cell to 'foo'" do
      @sheet[0, 0] = 'foo'
      @sheet[0, 0].value.should eq 'foo'
    end

    describe '#each' do
      it "should sort line in order of column" do
        @sheet.each_with_index do |cell, i|
          case i
          when 0
            cell.value.should eq 'simple'
          when 1
            cell.value.should eq 'workbook'
          when 2
            cell.value.should eq 'sheet1'
          when 3
            cell.value.should eq 'foo'
          when 4
            cell.value.should be_nil
          when 5
            cell.value.should eq 'foobaaa'
          end
        end
      end

      context "read sheet with blank" do
        include_context "sheet 'open book with blank'"

        it 'should get from ["A1"]' do
          @sheet_with_blank.each_with_index do |cell, i|
            case i
            when 5
              cell.value.should be_nil
            when 6
              cell.value.should eq 'simple'
            when 7
              cell.value.should be_nil
            when 8
              cell.value.should eq 'workbook'
            when 9
              cell.value.should eq 'sheet1'
            end
          end
        end
      end

    end

    describe "#each_row" do
      it "items should WrapExcel::Range" do
        @sheet.each_row do |rows|
          rows.should be_kind_of WrapExcel::Range
        end
      end

      context "with argument 1" do
        it 'should read from second row' do
          @sheet.each_row(1) do |rows|
            case rows.row
            when 2
              rows.values.should eq ['foo', nil, 'foobaaa']
            when 3
              rows.values.should eq ['matz', 'is', 'nice']
            end
          end
        end
      end

      context "read sheet with blank" do
        include_context "sheet 'open book with blank'"

        it 'should get from ["A1"]' do
          @sheet_with_blank.each_row do |rows|
            case rows.row - 1
            when 0
              rows.values.should eq [nil, nil, nil, nil, nil]
            when 1
              rows.values.should eq [nil, 'simple', nil, 'workbook', 'sheet1']
            when 2
              rows.values.should eq [nil, 'foo', nil, nil, 'foobaaa']
            when 3
              rows.values.should eq [nil, nil, nil, nil, nil]
            when 4
              rows.values.should eq [nil, 'matz', nil, 'is', 'nice']
            end
          end
        end
      end

    end

    describe "#each_row_with_index" do
      it "should read with index" do
        @sheet.each_row_with_index do |rows, idx|
          case idx
          when 0
            rows.values.should eq ['simple', 'workbook', 'sheet1']
          when 1
            rows.values.should eq ['foo', nil, 'foobaaa']
          when 2
            rows.values.should eq ['matz', 'is', 'nice']
          end
        end
      end

      context "with argument 1" do
        it "should read from second row, index is started 0" do
          @sheet.each_row_with_index(1) do |rows, idx|
            case idx
            when 0
              rows.values.should eq ['foo', nil, 'foobaaa']
            when 1
              rows.values.should eq ['matz', 'is', 'nice']
            end
          end
        end
      end

    end

    describe "#each_column" do
      it "items should WrapExcel::Range" do
        @sheet.each_column do |columns|
          columns.should be_kind_of WrapExcel::Range
        end
      end

      context "with argument 1" do
        it "should read from second column" do
          @sheet.each_column(1) do |columns|
            case columns.column
            when 2
              columns.values.should eq ['workbook', nil, 'is']
            when 3
              columns.values.should eq ['sheet1', 'foobaaa', 'nice']
            end
          end
        end
      end

      context "read sheet with blank" do
        include_context "sheet 'open book with blank'"

        it 'should get from ["A1"]' do
          @sheet_with_blank.each_column do |columns|
            case columns.column- 1
            when 0
              columns.values.should eq [nil, nil, nil, nil, nil]
            when 1
              columns.values.should eq [nil, 'simple', 'foo', nil, 'matz']
            when 2
              columns.values.should eq [nil, nil, nil, nil, nil]
            when 3
              columns.values.should eq [nil, 'workbook', nil, nil, 'is']
            when 4
              columns.values.should eq [nil, 'sheet1', 'foobaaa', nil, 'nice']
            end
          end
        end
      end

      context "read sheet which last cell is merged" do
        before do
          @book_merge_cells = WrapExcel::Book.open(@dir + '/merge_cells.xls')
          @sheet_merge_cell = @book_merge_cells[0]
        end

        after do
          @book_merge_cells.close
        end

        it "should get from ['A1'] to ['C2']" do
          columns_values = []
          @sheet_merge_cell.each_column do |columns|
            columns_values << columns.values
          end
          columns_values.should eq [
                                [nil, 'first merged', nil, 'merged'],
                                [nil, 'first merged', 'first', 'merged'],
                                [nil, 'first merged', 'second', 'merged'],
                                [nil, nil, 'third', 'merged']
                           ]
        end

      end

    end

    describe "#each_column_with_index" do
      it "should read with index" do
        @sheet.each_column_with_index do |columns, idx|
          case idx
          when 0
            columns.values.should eq ['simple', 'foo', 'matz']
          when 1
            columns.values.should eq ['workbook', nil, 'is']
          when 2
            columns.values.should eq ['sheet1', 'foobaaa', 'nice']
          end
        end
      end

      context "with argument 1" do
        it "should read from second column, index is started 0" do
          @sheet.each_column_with_index(1) do |column_range, idx|
            case idx
            when 0
              column_range.values.should eq ['workbook', nil, 'is']
            when 1
              column_range.values.should eq ['sheet1', 'foobaaa', 'nice']
            end
          end
        end
      end
    end

    describe "#row_range" do
      context "with second argument" do
        before do
          @row_range = @sheet.row_range(0, 1..2)
        end

        it { @row_range.should be_kind_of WrapExcel::Range }

        it "should get range cells of second argument" do
          @row_range.values.should eq ['workbook', 'sheet1']
        end
      end

      context "without second argument" do
        before do
          @row_range = @sheet.row_range(2)
        end

        it "should get all cells" do
          @row_range.values.should eq ['matz', 'is', 'nice']
        end
      end

    end

    describe "#col_range" do
      context "with second argument" do
        before do
          @col_range = @sheet.col_range(0, 1..2)
        end

        it { @col_range.should be_kind_of WrapExcel::Range }

        it "should get range cells of second argument" do
          @col_range.values.should eq ['foo', 'matz']
        end
      end

      context "without second argument" do
        before do
          @col_range = @sheet.col_range(1)
        end

        it "should get all cells" do
          @col_range.values.should eq ['workbook', nil, 'is']
        end
      end

    end

    describe "#method_missing" do
      it "can access COM method" do
        @sheet.Cells(1,1).Value.should eq 'simple'
      end

      context "unknown method" do
        it { expect { @sheet.hogehogefoo }.to raise_error }
      end
    end

  end
end
