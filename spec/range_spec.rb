# -*- coding: utf-8 -*-
require File.join(File.dirname(__FILE__), './spec_helper')

describe WrapExcel::Range do
  before do
    @dir = create_tmpdir
    @book = WrapExcel::Book.open(@dir + '/simple.xls')
    @sheet = @book[1]
    @range = WrapExcel::Range.new(@sheet.sheet.UsedRange.Rows(1))
  end

  after do
    @book.close
    rm_tmp(@dir)
  end

  describe "#each" do
    it "items is WrapExcel::Cell" do
      @range.each do |cell|
        cell.should be_kind_of WrapExcel::Cell
      end
    end
  end

  describe "#values" do
    context "with (0..2)" do
      it { @range.values(0..2).should eq ['simple', 'file', 'sheet2'] }
    end

    context "with (1..2)" do
      it { @range.values(1..2).should eq ['file', 'sheet2'] }
    end

    context "with (2..2)" do
      it { @range.values(2..2).should eq ['sheet2'] }
    end

    context "with no arguments" do
      it { @range.values.should eq ['simple', 'file', 'sheet2'] }
    end

    context "when instance is column range" do
      before do
        @sheet = @book[0]
        @range = WrapExcel::Range.new(@sheet.sheet.UsedRange.Columns(1))
      end
      it { @range.values.should eq ['simple', 'foo', 'matz'] }
    end
  end

  describe "#[]" do
    context "access [0]" do
      it { @range[0].should be_kind_of WrapExcel::Cell }
      it { @range[0].value.should eq 'simple' }
    end

    context "access [2]" do
      it { @range[2].value.should eq 'sheet2' }
    end

    context "access [0] and [1] and [2]" do
      it "should get every values" do
        @range[0].value.should eq 'simple'
        @range[1].value.should eq 'file'
        @range[2].value.should eq 'sheet2'
      end
    end
  end

  describe "#method_missing" do
    it "can access COM method" do
      @range.Range(@range.Cells.Item(1), @range.Cells.Item(3)).value.should eq [@range.values(0..2)]
    end

    context "unknown method" do
      it { expect { @range.hogehogefoo}.to raise_error }
    end
  end
end
